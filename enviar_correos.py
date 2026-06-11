#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Automatizador de envío U4U por SMTP.
- Envía por lotes los registros con estado Pendiente.
- Marca como Enviado, Error o deja Pendiente los faltantes.
- Personaliza el saludo de la plantilla HTML.
- Permite prueba previa y BCC aleatorio por lote.

Uso típico:
  python enviar_correos.py --dry-run
  python enviar_correos.py --test-to emmanuelquintan2020@gmail.com --yes
  python enviar_correos.py --limit 100 --yes --enable-random-bcc
"""

from __future__ import annotations

import argparse
import csv
import html as html_lib
import io
import json
import mimetypes
import os
import random
import re
import smtplib
import sys
import time
from datetime import datetime
from email.message import EmailMessage
from email.utils import formataddr, make_msgid
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

ROOT = Path(__file__).resolve().parent
DEFAULT_CSV = ROOT / "contactos.csv"
DEFAULT_TEMPLATE = ROOT / "plantilla_u4u.html"
DEFAULT_ATTACHMENTS_DIR = ROOT / "adjuntos"
DEFAULT_LOG_DIR = ROOT / "logs"
DEFAULT_LOG_DIR.mkdir(exist_ok=True)

EMAIL_RE = re.compile(r"^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", re.I)
PLACEHOLDER_RE = re.compile(r"\[[^\[\]\r\n]{1,120}\]")
PENDING_STATES = {"", "pendiente"}
SKIP_STATES = {"enviado", "baja", "cancelado", "revisar", "revisar datos"}

REQUIRED_COLUMNS = [
    "id", "empresa", "contacto", "destinatarios", "estado", "fecha_envio",
    "ultimo_intento", "intentos", "ultimo_error", "bcc_aplicado", "comentarios"
]


def log(msg: str) -> None:
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line, flush=True)
    with (DEFAULT_LOG_DIR / "envios.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_env(path: Path = ROOT / ".env") -> Dict[str, str]:
    """Carga un .env simple sin dependencias externas."""
    env = dict(os.environ)
    if path.exists():
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            env[key] = value
    return env


def save_env(values: Dict[str, str], path: Path = ROOT / ".env") -> None:
    """Guarda configuración local en .env. Este archivo no debe compartirse."""
    ordered_keys = [
        "SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASSWORD", "FROM_NAME", "REPLY_TO",
        "SUBJECT", "DEFAULT_LIMIT", "SLEEP_SECONDS", "SLEEP_JITTER_SECONDS", "TEST_EMAIL",
        "TEMPLATE_PATH", "PLACEHOLDER_MAP", "BCC_RANDOM_ENABLED", "BCC_RANDOM_PER_100", "BCC_RECIPIENTS",
        "ATTACHMENTS_ENABLED",
    ]
    lines = [
        "# Configuración local U4U. No compartas este archivo.",
        "# La contraseña SMTP se guarda solo en tu computadora.",
        "",
    ]
    for key in ordered_keys:
        if key in values:
            val = str(values.get(key, "")).replace("\n", " ").strip()
            lines.append(f"{key}={val}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_bool(value: str, default: bool = False) -> bool:
    if value is None:
        return default
    value = str(value).strip().lower()
    if value in {"1", "true", "yes", "si", "sí", "on", "y"}:
        return True
    if value in {"0", "false", "no", "off", "n"}:
        return False
    return default


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sniff_csv(text: str) -> Tuple[str, str]:
    """Devuelve (texto_sin_sep, delimitador). Soporta línea Excel 'sep=;'"""
    lines = text.splitlines()
    if lines and lines[0].lower().startswith("sep="):
        delim = lines[0][4:5] or ";"
        return "\n".join(lines[1:]), delim
    try:
        dialect = csv.Sniffer().sniff(text[:2048], delimiters=";,\t")
        return text, dialect.delimiter
    except csv.Error:
        return text, ";"


def read_contacts(csv_path: Path) -> Tuple[List[Dict[str, str]], str]:
    if not csv_path.exists():
        raise FileNotFoundError(f"No encontré el archivo de contactos: {csv_path}")
    text = csv_path.read_text(encoding="utf-8-sig")
    text, delimiter = sniff_csv(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    rows = []
    for row in reader:
        clean = {str(k).strip(): (v or "").strip() for k, v in row.items() if k is not None}
        if not any(clean.values()):
            continue
        for col in REQUIRED_COLUMNS:
            clean.setdefault(col, "")
        rows.append(clean)
    return rows, delimiter


def write_contacts(csv_path: Path, rows: List[Dict[str, str]]) -> None:
    import stat
    tmp = csv_path.with_suffix(".tmp")
    
    # Asegurar que no quede como solo lectura si el archivo ya existe
    try:
        if csv_path.exists():
            mode = csv_path.stat().st_mode
            if not (mode & stat.S_IWRITE):
                csv_path.chmod(stat.S_IWRITE)
    except Exception:
        pass

    # Escribir en el archivo temporal
    try:
        with tmp.open("w", encoding="utf-8-sig", newline="") as f:
            f.write("sep=;\n")
            writer = csv.DictWriter(f, fieldnames=REQUIRED_COLUMNS, delimiter=";", extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as exc:
        log(f"Error al escribir en archivo temporal {tmp.name}: {exc}")
        raise exc

    # Reemplazar con reintentos si el archivo principal está bloqueado (ej. por Excel)
    max_retries = 40  # 40 intentos * 3 seg = 120 segundos (2 minutos)
    for attempt in range(1, max_retries + 1):
        try:
            tmp.replace(csv_path)
            return
        except PermissionError as exc:
            # WinError 5 es un PermissionError en Python
            if attempt == 1:
                log(f"[ADVERTENCIA] No se pudo guardar '{csv_path.name}'. El archivo podría estar abierto en Excel o bloqueado por otro proceso.")
                log("Por favor, CIERRA Excel o el programa que tenga abierto el archivo para guardar el estado y continuar...")
            elif attempt % 5 == 0:
                log(f"Reintento {attempt}/{max_retries} para guardar '{csv_path.name}'. Asegúrate de cerrar el archivo en Excel.")
            time.sleep(3)
        except Exception as exc:
            # Si ocurre otro tipo de error, intentamos limpiar el archivo temporal y relanzamos
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass
            log(f"Error inesperado al reemplazar '{csv_path.name}': {exc}")
            raise exc

    # Si se agotan todos los reintentos, limpiamos y levantamos el error
    try:
        tmp.unlink(missing_ok=True)
    except Exception:
        pass
    msg = f"No se pudo reemplazar '{csv_path.name}' después de {max_retries} intentos. Asegúrate de cerrar Excel u otros programas que lo usen."
    log(f"[ERROR CRÍTICO] {msg}")
    raise PermissionError(msg)


def parse_recipients(value: str) -> List[str]:
    parts = re.split(r"[,;\s]+", value or "")
    recipients = []
    for part in parts:
        email = part.strip().strip("<>").lower()
        if email and EMAIL_RE.match(email) and email not in recipients:
            recipients.append(email)
    return recipients


def limpiar_contacto(contacto: str) -> str:
    contacto = re.sub(r"\s+", " ", (contacto or "").strip())
    if not contacto:
        return ""
    # Soporta contactos tipo "Karen / Dr. José" o "Ruben, Guillermo".
    parts = [p.strip(" .") for p in re.split(r"\s*/\s*|\s*;\s*|\s*,\s*", contacto) if p.strip(" .")]
    if len(parts) == 0:
        return ""
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return f"{parts[0]} y {parts[1]}"
    return ", ".join(parts[:-1]) + f" y {parts[-1]}"


def detect_placeholders(template_html: str) -> List[str]:
    return sorted(set(PLACEHOLDER_RE.findall(template_html)), key=str.lower)


def normalize_placeholder_label(placeholder: str) -> str:
    label = placeholder.strip().strip("[]").lower()
    return re.sub(r"[^a-z0-9áéíóúüñ]+", " ", label).strip()


def default_column_for_placeholder(placeholder: str, columns: Iterable[str]) -> str:
    available = [col for col in columns if col]
    lower_map = {col.lower(): col for col in available}
    label = normalize_placeholder_label(placeholder)
    rules = [
        (("nombre", "contacto", "compras", "persona"), "contacto"),
        (("empresa", "institucion", "institución", "clinica", "clínica", "hospital"), "empresa"),
        (("correo", "email", "mail", "destinatario"), "destinatarios"),
        (("id", "folio"), "id"),
        (("comentario", "nota"), "comentarios"),
    ]
    for needles, column in rules:
        if any(needle in label for needle in needles) and column in lower_map:
            return lower_map[column]
    if "contacto" in lower_map:
        return lower_map["contacto"]
    return available[0] if available else ""


def load_placeholder_map(env: Dict[str, str], placeholders: Iterable[str] = (), columns: Iterable[str] = REQUIRED_COLUMNS) -> Dict[str, str]:
    raw = env.get("PLACEHOLDER_MAP", "").strip()
    result: Dict[str, str] = {}
    if raw:
        try:
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                result = {str(key): str(value) for key, value in parsed.items() if str(value).strip()}
        except json.JSONDecodeError:
            for part in raw.split(";"):
                if "=" in part:
                    key, value = part.split("=", 1)
                    if key.strip() and value.strip():
                        result[key.strip()] = value.strip()
    for placeholder in placeholders:
        result.setdefault(placeholder, default_column_for_placeholder(placeholder, columns))
    return result


def dump_placeholder_map(mapping: Dict[str, str]) -> str:
    clean = {str(key): str(value) for key, value in mapping.items() if str(key).strip() and str(value).strip()}
    return json.dumps(clean, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def value_for_placeholder(row: Dict[str, str], column: str) -> str:
    if not column:
        return ""
    if column.lower() == "contacto":
        return limpiar_contacto(row.get(column, ""))
    return row.get(column, "")


def apply_placeholder_map(template_html: str, row: Dict[str, str], env: Dict[str, str]) -> str:
    placeholders = detect_placeholders(template_html)
    columns = list(row.keys()) or REQUIRED_COLUMNS
    mapping = load_placeholder_map(env, placeholders, columns)
    html = template_html
    for placeholder in placeholders:
        column = mapping.get(placeholder, "")
        value = value_for_placeholder(row, column)
        html = html.replace(placeholder, html_lib.escape(value, quote=True))
    return html


def personalize_template(template_html: str, row: Dict[str, str], env: Optional[Dict[str, str]] = None) -> str:
    env = env or {}
    contacto = limpiar_contacto(row.get("contacto", ""))
    if contacto:
        saludo = f"¡Hola <strong>{html_lib.escape(contacto)}</strong>! 👋"
        fallback_nombre = html_lib.escape(contacto)
    else:
        saludo = "¡Hola! 👋"
        fallback_nombre = ""

    html = apply_placeholder_map(template_html, row, env)
    # Reemplazo principal para la plantilla actual.
    pattern = r"¡Hola\s*<strong>\s*\[Nombre\s*/\s*Equipo de compras\]\s*</strong>\s*!\s*👋?"
    html, _ = re.subn(pattern, saludo, html, flags=re.I)

    # Reemplazos de respaldo por si cambias ligeramente la plantilla.
    html = html.replace("[Nombre / Equipo de compras]", fallback_nombre)
    html = html.replace("[NOMBRE / EQUIPO DE COMPRAS]", fallback_nombre)
    html = html.replace("[Contacto]", fallback_nombre)
    html = html.replace("[CONTACTO]", fallback_nombre)
    html = html.replace("[Empresa]", html_lib.escape(row.get("empresa", "")))
    html = html.replace("[EMPRESA]", html_lib.escape(row.get("empresa", "")))

    # Si no había contacto, elimina saludos vacíos con strong.
    html = re.sub(r"¡Hola\s*<strong>\s*</strong>\s*!", "¡Hola!", html, flags=re.I)
    return html


def html_to_plain_text(html: str) -> str:
    text = re.sub(r"<style[\s\S]*?</style>", " ", html, flags=re.I)
    text = re.sub(r"<script[\s\S]*?</script>", " ", text, flags=re.I)
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"</p>|</tr>|</h[1-6]>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html_lib.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def collect_attachments(folder: Path) -> List[Path]:
    if not folder.exists():
        return []
    attachments = []
    ignored = {"README_ADJUNTOS.txt", ".gitkeep"}
    for path in sorted(folder.iterdir()):
        if path.is_file() and not path.name.startswith(".") and path.name not in ignored:
            attachments.append(path)
    return attachments


def build_message(
    row: Dict[str, str],
    recipients: List[str],
    template_html: str,
    env: Dict[str, str],
    attachments: List[Path],
    bcc_recipients: Optional[List[str]] = None,
) -> EmailMessage:
    html = personalize_template(template_html, row, env)
    text = html_to_plain_text(html)

    subject = env.get("SUBJECT", "Catálogo U4U para depósitos dentales")
    smtp_user = env.get("SMTP_USER", "")
    from_name = env.get("FROM_NAME", "Auri Morales - U4U")
    reply_to = env.get("REPLY_TO", smtp_user)

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr((from_name, smtp_user))
    msg["To"] = ", ".join(recipients)
    if bcc_recipients:
        msg["Bcc"] = ", ".join(bcc_recipients)
    if reply_to:
        msg["Reply-To"] = reply_to
        msg["List-Unsubscribe"] = f"<mailto:{reply_to}?subject=BAJA>"
    msg["Message-ID"] = make_msgid(domain=(smtp_user.split("@")[-1] if "@" in smtp_user else None))
    msg.set_content(text, subtype="plain", charset="utf-8")
    msg.add_alternative(html, subtype="html", charset="utf-8")

    for attachment in attachments:
        ctype, encoding = mimetypes.guess_type(str(attachment))
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(
            attachment.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=attachment.name,
        )
    return msg


def select_pending(rows: List[Dict[str, str]], limit: int, retry_errors: bool = False) -> List[int]:
    selected = []
    for i, row in enumerate(rows):
        state = (row.get("estado") or "").strip().lower()
        if state in PENDING_STATES or (retry_errors and state == "error"):
            selected.append(i)
        if len(selected) >= limit:
            break
    return selected


def remaining_pending(rows: List[Dict[str, str]]) -> int:
    total = 0
    for row in rows:
        state = (row.get("estado") or "").strip().lower()
        if state in PENDING_STATES:
            total += 1
    return total


def count_by_state(rows: List[Dict[str, str]]) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for row in rows:
        state = (row.get("estado") or "Pendiente").strip() or "Pendiente"
        result[state] = result.get(state, 0) + 1
    return result


def normalize_contact_state(value: str) -> str:
    state = (value or "").strip() or "Pendiente"
    if state.lower() in {"", "pendiente"}:
        return "Pendiente"
    return state


def update_contact_status(rows: List[Dict[str, str]], index: int, value: str) -> None:
    if 0 <= index < len(rows):
        rows[index]["estado"] = normalize_contact_state(value)


def bulk_update_contact_status(rows: List[Dict[str, str]], value: str) -> None:
    normalized = normalize_contact_state(value)
    for row in rows:
        row["estado"] = normalized


def require_env(env: Dict[str, str]) -> None:
    required = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASSWORD"]
    missing = [k for k in required if not env.get(k) or env.get(k) == "PEGA_AQUI_TU_CONTRASENA_SMTP"]
    if missing:
        raise RuntimeError(
            "Faltan estos datos en .env: " + ", ".join(missing) +
            "\nAbre la interfaz, guarda la configuración y pega tu contraseña SMTP."
        )


def choose_random_bcc_indexes(selected_indexes: List[int], per_100: int) -> Set[int]:
    if per_100 <= 0 or not selected_indexes:
        return set()
    amount = round(len(selected_indexes) * per_100 / 100)
    amount = max(0, min(len(selected_indexes), amount))
    return set(random.sample(selected_indexes, amount)) if amount else set()


def get_bcc_config(env: Dict[str, str], args: argparse.Namespace) -> Tuple[bool, int, List[str]]:
    enabled = args.enable_random_bcc or parse_bool(env.get("BCC_RANDOM_ENABLED", "true"), True)
    per_100_raw = args.bcc_random_per_100 if args.bcc_random_per_100 is not None else env.get("BCC_RANDOM_PER_100", "10")
    try:
        per_100 = int(per_100_raw)
    except (TypeError, ValueError):
        per_100 = 10
    per_100 = max(0, min(100, per_100))
    recipients_raw = args.bcc_recipients or env.get("BCC_RECIPIENTS", "bvilenski@gmail.com, arturogoldsmit@gmail.com")
    recipients = parse_recipients(recipients_raw)
    if enabled and not recipients:
        raise RuntimeError("BCC aleatorio está activo, pero BCC_RECIPIENTS está vacío o no tiene correos válidos.")
    return enabled, per_100, recipients


def main() -> int:
    parser = argparse.ArgumentParser(description="Enviar correos U4U por lotes con control de pendientes.")
    parser.add_argument("--csv", default=str(DEFAULT_CSV), help="Ruta del CSV de contactos.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Ruta de la plantilla HTML.")
    parser.add_argument("--adjuntos", default=str(DEFAULT_ATTACHMENTS_DIR), help="Carpeta de adjuntos.")
    parser.add_argument("--limit", type=int, default=None, help="Máximo de registros a enviar en esta corrida.")
    parser.add_argument("--dry-run", action="store_true", help="Solo muestra lo que enviaría; no manda correos ni cambia estados.")
    parser.add_argument("--test-to", default="", help="Envía una prueba a este correo sin cambiar estados.")
    parser.add_argument("--retry-errors", action="store_true", help="Reintenta registros con estado Error además de Pendiente.")
    parser.add_argument("--yes", action="store_true", help="Confirma el envío sin preguntar.")
    parser.add_argument("--enable-random-bcc", action="store_true", help="Activa BCC aleatorio para una parte del lote.")
    parser.add_argument("--bcc-random-per-100", type=int, default=None, help="Cantidad aproximada de BCC por cada 100 registros. Ej: 10.")
    parser.add_argument("--bcc-recipients", default="", help="Correos BCC separados por coma.")
    args = parser.parse_args()

    env = load_env()
    limit = args.limit or int(env.get("DEFAULT_LIMIT", "100"))
    sleep_seconds = float(env.get("SLEEP_SECONDS", "5"))
    sleep_jitter_seconds = float(env.get("SLEEP_JITTER_SECONDS", "5"))

    csv_path = Path(args.csv)
    template_path = Path(args.template)
    attachments_dir = Path(args.adjuntos)

    rows, _ = read_contacts(csv_path)
    template_html = template_path.read_text(encoding="utf-8")
    attachments_enabled = parse_bool(env.get("ATTACHMENTS_ENABLED", "true"), True)
    attachments = collect_attachments(attachments_dir) if attachments_enabled else []

    selected_indexes = select_pending(rows, limit, retry_errors=args.retry_errors)
    bcc_enabled, bcc_per_100, bcc_recipients = get_bcc_config(env, args)
    random_bcc_indexes = choose_random_bcc_indexes(selected_indexes, bcc_per_100) if bcc_enabled and not args.test_to else set()

    if args.dry_run:
        log(f"DRY-RUN: registros pendientes seleccionados: {len(selected_indexes)} de {remaining_pending(rows)} pendientes totales.")
        if bcc_enabled:
            log(f"DRY-RUN: BCC aleatorio activo: {len(random_bcc_indexes)} registros del lote llevarían BCC a {', '.join(bcc_recipients)}.")
        for n, idx in enumerate(selected_indexes[:10], start=1):
            row = rows[idx]
            mark = " [BCC]" if idx in random_bcc_indexes else ""
            log(f"{n}. {row.get('empresa','')} | {row.get('contacto','')} | {row.get('destinatarios','')}{mark}")
        if len(selected_indexes) > 10:
            log(f"... y {len(selected_indexes)-10} más en esta corrida.")
        return 0

    require_env(env)

    # Prueba: no toca contactos.csv y no aplica BCC aleatorio.
    if args.test_to:
        recipients = parse_recipients(args.test_to)
        if not recipients:
            raise RuntimeError("El correo de prueba no parece válido.")
        sample_row = dict(rows[selected_indexes[0]]) if selected_indexes else {
            "empresa": "", "contacto": "", "destinatarios": args.test_to,
            "estado": "Pendiente", "fecha_envio": "", "ultimo_intento": "", "intentos": "0", "ultimo_error": "", "bcc_aplicado": "", "comentarios": ""
        }
        sample_row["contacto"] = "JOSE EMMANUEL QUINTANA TORRES"
        sample_row["destinatarios"] = args.test_to
        msg = build_message(sample_row, recipients, template_html, env, attachments)
        host = env["SMTP_HOST"]
        port = int(env.get("SMTP_PORT", "465"))
        log(f"Enviando prueba a {', '.join(recipients)}...")
        with smtplib.SMTP_SSL(host, port, timeout=60) as smtp:
            smtp.login(env["SMTP_USER"], env["SMTP_PASSWORD"])
            smtp.send_message(msg)
        log("Prueba enviada correctamente. No se modificó contactos.csv.")
        return 0

    if not selected_indexes:
        log("No hay registros pendientes para enviar.")
        return 0

    log(
        f"Se enviarán hasta {len(selected_indexes)} registros. Adjuntos {'habilitados' if attachments_enabled else 'deshabilitados'}: {len(attachments)} detectados."
    )
    if bcc_enabled:
        log(f"BCC aleatorio activo: {len(random_bcc_indexes)} registros del lote llevarán BCC a {', '.join(bcc_recipients)}.")
    if not args.yes:
        answer = input("Escribe ENVIAR para confirmar: ").strip().upper()
        if answer != "ENVIAR":
            log("Operación cancelada por el usuario.")
            return 1

    host = env["SMTP_HOST"]
    port = int(env.get("SMTP_PORT", "465"))
    sent = 0
    errors = 0

    with smtplib.SMTP_SSL(host, port, timeout=60) as smtp:
        smtp.login(env["SMTP_USER"], env["SMTP_PASSWORD"])
        for idx in selected_indexes:
            row = rows[idx]
            recipients = parse_recipients(row.get("destinatarios", ""))
            row["ultimo_intento"] = now_iso()
            try:
                row["intentos"] = str(int(row.get("intentos") or "0") + 1)
            except ValueError:
                row["intentos"] = "1"

            if not recipients:
                row["estado"] = "Revisar datos"
                row["ultimo_error"] = "Sin destinatarios válidos"
                write_contacts(csv_path, rows)
                errors += 1
                log(f"REVISAR: {row.get('empresa','')} sin destinatarios válidos.")
                continue

            bcc_for_this = bcc_recipients if idx in random_bcc_indexes else []
            try:
                msg = build_message(row, recipients, template_html, env, attachments, bcc_for_this)
                smtp.send_message(msg)
                row["estado"] = "Enviado"
                row["fecha_envio"] = now_iso()
                row["ultimo_error"] = ""
                row["bcc_aplicado"] = ", ".join(bcc_for_this) if bcc_for_this else ""
                sent += 1
                bcc_label = f" | BCC: {', '.join(bcc_for_this)}" if bcc_for_this else ""
                log(f"ENVIADO: {row.get('empresa','')} -> {', '.join(recipients)}{bcc_label}")
            except Exception as exc:
                row["estado"] = "Error"
                row["ultimo_error"] = str(exc)[:500]
                errors += 1
                log(f"ERROR: {row.get('empresa','')} -> {', '.join(recipients)} | {exc}")
            finally:
                write_contacts(csv_path, rows)

            pause = sleep_seconds + (random.uniform(0, sleep_jitter_seconds) if sleep_jitter_seconds > 0 else 0)
            if pause > 0:
                time.sleep(pause)

    log(f"Resumen: enviados={sent}, errores={errors}, pendientes_restantes={remaining_pending(rows)}")
    return 0 if errors == 0 else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        log("Interrumpido por el usuario. Los envíos ya completados quedaron guardados.")
        raise SystemExit(130)
    except Exception as exc:
        log(f"FALLÓ: {exc}")
        raise SystemExit(1)
