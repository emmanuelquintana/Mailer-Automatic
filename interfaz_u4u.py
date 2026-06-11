#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz gráfica moderna para el automatizador de correos U4U.

Flujo recomendado:
1) Configura SMTP y guarda.
2) Revisa o carga la plantilla activa.
3) Envía 1 correo de prueba.
4) Aprueba la prueba.
5) Envía el batch de pendientes.
"""

from __future__ import annotations

import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import html as html_lib
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import enviar_correos as core

ROOT = Path(__file__).resolve().parent
PYTHON = sys.executable
DEFAULT_TEMPLATE = ROOT / "plantilla_u4u.html"

COLORS = {
    "app": "#f4f7f8",
    "panel": "#ffffff",
    "sidebar": "#0f2f36",
    "sidebar_soft": "#153d45",
    "ink": "#172a2f",
    "muted": "#66777d",
    "line": "#d9e4e7",
    "accent": "#1f8f8b",
    "accent_dark": "#126c6a",
    "danger": "#a23b3b",
    "warning": "#a76a19",
    "success": "#1d7c58",
    "terminal": "#101820",
}

FONT_UI = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)
FONT_TITLE = ("Segoe UI Semibold", 22)
FONT_SECTION = ("Segoe UI Semibold", 12)
FONT_MONO = ("Consolas", 10)


class U4UMailerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("U4U Mailer Studio")
        self.geometry("1240x820")
        self.minsize(1080, 700)
        self.configure(bg=COLORS["app"])

        self.queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = False
        self.test_sent_ok = False
        self.test_approved = False
        self.template_dirty = False
        self.current_page = "panel"
        self.preview_html = ""
        self._preview_render_width = 0
        self._preview_resize_job: str | None = None
        self._preview_render_job: str | None = None
        self._preview_request_id = 0
        self.preview_image: tk.PhotoImage | None = None

        self.vars: dict[str, tk.StringVar] = {}
        self.metric_vars: dict[str, tk.StringVar] = {}
        self.placeholder_vars: dict[str, tk.StringVar] = {}
        self.placeholder_rows: dict[str, tk.Frame] = {}
        self.placeholder_map: dict[str, str] = {}
        self.csv_columns = list(core.REQUIRED_COLUMNS)
        self.nav_buttons: dict[str, tk.Button] = {}
        self.pages: dict[str, tk.Frame] = {}
        self.bcc_enabled = tk.BooleanVar(value=True)
        self.attachments_enabled = tk.BooleanVar(value=True)

        self._build_style()
        self._build_ui()
        self._load_config()
        self._load_template_editor()
        self.refresh_counts()
        self.show_page("panel")
        self.after(150, self._process_queue)

    def _build_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Modern.TEntry", padding=8, relief="flat", fieldbackground="#ffffff", foreground=COLORS["ink"])
        style.configure("Modern.TCombobox", padding=8, relief="flat", fieldbackground="#ffffff", foreground=COLORS["ink"])
        style.configure("Modern.TCheckbutton", background=COLORS["panel"], foreground=COLORS["ink"], font=FONT_UI)
        style.configure("Clean.Vertical.TScrollbar", background=COLORS["line"], troughcolor=COLORS["app"], bordercolor=COLORS["app"], arrowcolor=COLORS["muted"])

    def _build_ui(self) -> None:
        shell = tk.Frame(self, bg=COLORS["app"])
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        self.sidebar = tk.Frame(shell, bg=COLORS["sidebar"], width=248)
        self.sidebar.grid(row=0, column=0, sticky="ns")
        self.sidebar.grid_propagate(False)
        self._build_sidebar()

        content = tk.Frame(shell, bg=COLORS["app"])
        content.grid(row=0, column=1, sticky="nsew")
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=1)

        self.header = tk.Frame(content, bg=COLORS["app"])
        self.header.grid(row=0, column=0, sticky="ew", padx=28, pady=(24, 12))
        self.header.columnconfigure(0, weight=1)
        self.page_title = tk.Label(self.header, text="", bg=COLORS["app"], fg=COLORS["ink"], font=FONT_TITLE)
        self.page_title.grid(row=0, column=0, sticky="w")
        self.page_subtitle = tk.Label(self.header, text="", bg=COLORS["app"], fg=COLORS["muted"], font=FONT_UI)
        self.page_subtitle.grid(row=1, column=0, sticky="w", pady=(3, 0))
        self.status_pill = tk.Label(
            self.header,
            text="Listo",
            bg="#e7f4f2",
            fg=COLORS["accent_dark"],
            font=("Segoe UI Semibold", 9),
            padx=14,
            pady=6,
        )
        self.status_pill.grid(row=0, column=1, rowspan=2, sticky="e")

        self.page_host = tk.Frame(content, bg=COLORS["app"])
        self.page_host.grid(row=1, column=0, sticky="nsew", padx=28, pady=(0, 24))
        self.page_host.rowconfigure(0, weight=1)
        self.page_host.columnconfigure(0, weight=1)

        self._build_dashboard_page()
        self._build_template_page()
        self._build_config_page()
        self._build_send_page()
        self._build_contacts_page()
        self._build_log_page()

    def _build_sidebar(self) -> None:
        brand = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        brand.pack(fill="x", padx=22, pady=(24, 28))
        tk.Label(brand, text="U4U", bg=COLORS["sidebar"], fg="#f4fbfb", font=("Segoe UI Semibold", 25)).pack(anchor="w")
        tk.Label(brand, text="Mailer Studio", bg=COLORS["sidebar"], fg="#b9d2d4", font=FONT_UI).pack(anchor="w")

        nav_items = [
            ("panel", "Panel"),
            ("plantilla", "Plantilla"),
            ("config", "SMTP"),
            ("envio", "Prueba y envio"),
            ("contactos", "Contactos"),
            ("bitacora", "Bitacora"),
        ]
        for key, label in nav_items:
            btn = tk.Button(
                self.sidebar,
                text=label,
                command=lambda name=key: self.show_page(name),
                anchor="w",
                padx=18,
                pady=12,
                bd=0,
                relief="flat",
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
                bg=COLORS["sidebar"],
                fg="#d7e7e8",
                activebackground=COLORS["sidebar_soft"],
                activeforeground="#ffffff",
            )
            btn.pack(fill="x", padx=14, pady=2)
            self.nav_buttons[key] = btn

        footer = tk.Frame(self.sidebar, bg=COLORS["sidebar"])
        footer.pack(side="bottom", fill="x", padx=22, pady=22)
        tk.Label(footer, text="Plantilla activa", bg=COLORS["sidebar"], fg="#8fb6ba", font=FONT_SMALL).pack(anchor="w")
        self.sidebar_template = tk.Label(
            footer,
            text="plantilla_u4u.html",
            bg=COLORS["sidebar"],
            fg="#ffffff",
            font=("Segoe UI Semibold", 9),
            wraplength=190,
            justify="left",
        )
        self.sidebar_template.pack(anchor="w", pady=(4, 0))

    def _build_dashboard_page(self) -> None:
        page = self._new_page("panel")
        page.columnconfigure(0, weight=2)
        page.columnconfigure(1, weight=1)
        page.rowconfigure(2, weight=1)

        metrics = tk.Frame(page, bg=COLORS["app"])
        metrics.grid(row=0, column=0, columnspan=2, sticky="ew")
        for i in range(4):
            metrics.columnconfigure(i, weight=1)
        self._metric_card(metrics, "Pendientes", "pending", 0, COLORS["accent"])
        self._metric_card(metrics, "Proximo batch", "batch", 1, COLORS["ink"])
        self._metric_card(metrics, "BCC estimado", "bcc", 2, COLORS["warning"])
        self._metric_card(metrics, "Plantilla", "template", 3, COLORS["success"])

        flow = self._surface(page, "Flujo de trabajo", "Controla el envio desde una secuencia clara y reversible.")
        flow.grid(row=1, column=0, sticky="nsew", pady=(18, 0), padx=(0, 10))
        flow.columnconfigure(0, weight=1)
        steps = [
            ("1", "Configura SMTP", "Guarda remitente, asunto, pausas y BCC."),
            ("2", "Revisa plantilla", "Visualiza, edita o carga otra plantilla HTML."),
            ("3", "Prueba", "Envia un correo de prueba antes del lote."),
            ("4", "Aprueba y envia", "Solo el batch aprobado toca contactos.csv."),
        ]
        for idx, (number, title, detail) in enumerate(steps):
            row = tk.Frame(flow, bg=COLORS["panel"])
            row.grid(row=idx + 1, column=0, sticky="ew", padx=2, pady=(12 if idx == 0 else 4, 4))
            row.columnconfigure(1, weight=1)
            badge = tk.Label(row, text=number, bg="#e8f4f3", fg=COLORS["accent_dark"], font=("Segoe UI Semibold", 10), width=3, pady=6)
            badge.grid(row=0, column=0, sticky="n", padx=(0, 12))
            tk.Label(row, text=title, bg=COLORS["panel"], fg=COLORS["ink"], font=FONT_SECTION).grid(row=0, column=1, sticky="w")
            tk.Label(row, text=detail, bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL).grid(row=1, column=1, sticky="w", pady=(2, 0))

        quick = self._surface(page, "Acciones rapidas", "Herramientas de revision sin salir de la app.")
        quick.grid(row=1, column=1, sticky="nsew", pady=(18, 0), padx=(10, 0))
        self._button(quick, "Actualizar conteos", self.refresh_counts, "secondary").grid(row=1, column=0, sticky="ew", pady=(12, 8))
        self._button(quick, "Abrir plantilla en navegador", self.open_template_preview, "secondary").grid(row=2, column=0, sticky="ew", pady=8)
        self._button(quick, "Abrir contactos.csv", lambda: self.open_path(ROOT / "contactos.csv"), "secondary").grid(row=3, column=0, sticky="ew", pady=8)
        self._button(quick, "Ir a prueba y envio", lambda: self.show_page("envio"), "primary").grid(row=4, column=0, sticky="ew", pady=(8, 0))

        activity = self._surface(page, "Actividad reciente", "Resumen operativo de esta sesion.")
        activity.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(18, 0))
        activity.rowconfigure(1, weight=1)
        activity.columnconfigure(0, weight=1)
        self.dashboard_log = tk.Text(
            activity,
            height=8,
            wrap="word",
            bd=0,
            bg="#f8fbfb",
            fg=COLORS["ink"],
            font=FONT_SMALL,
            padx=12,
            pady=10,
        )
        self.dashboard_log.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.dashboard_log.configure(state="disabled")

    def _build_template_page(self) -> None:
        page = self._new_page("plantilla")
        page.rowconfigure(2, weight=1)
        page.columnconfigure(0, weight=1)

        toolbar = self._surface(page, "Plantilla activa", "Visualiza, cambia, carga o edita el HTML que se usara en el envio.")
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(0, weight=1)
        self.vars["TEMPLATE_PATH"] = tk.StringVar()
        path_entry = ttk.Entry(toolbar, textvariable=self.vars["TEMPLATE_PATH"], style="Modern.TEntry")
        path_entry.grid(row=1, column=0, sticky="ew", pady=(14, 0), padx=(0, 10))
        self._button(toolbar, "Usar archivo", self.choose_template_file, "secondary").grid(row=1, column=1, padx=4, pady=(14, 0))
        self._button(toolbar, "Importar local", self.import_template_file, "secondary").grid(row=1, column=2, padx=4, pady=(14, 0))
        self._button(toolbar, "Vista navegador", self.open_template_preview, "primary").grid(row=1, column=3, padx=(4, 0), pady=(14, 0))

        mapping = self._surface(page, "Mapeo de placeholders", "Detecta campos [asi] y elige de que columna del CSV salen.")
        mapping.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        mapping.columnconfigure(0, weight=1)
        self.placeholder_map_frame = tk.Frame(mapping, bg=COLORS["panel"])
        self.placeholder_map_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        self.placeholder_map_frame.columnconfigure(1, weight=1)
        self.placeholder_status = tk.Label(
            mapping,
            text="Carga una plantilla para detectar placeholders.",
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=FONT_SMALL,
            anchor="w",
        )
        self.placeholder_status.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self._button(mapping, "Guardar mapeo", self.save_placeholder_mapping, "secondary").grid(row=1, column=1, sticky="ne", padx=(16, 0), pady=(12, 0))

        work = tk.Frame(page, bg=COLORS["app"])
        work.grid(row=2, column=0, sticky="nsew", pady=(18, 0))
        work.columnconfigure(0, weight=7)
        work.columnconfigure(1, weight=4)
        work.rowconfigure(0, weight=1)

        inspector = self._surface(work, "Vista visual")
        inspector.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        inspector.rowconfigure(1, weight=1)
        inspector.columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            inspector,
            bd=0,
            highlightthickness=0,
            bg="#eef4f5",
        )
        self.preview_canvas.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        preview_scroll = ttk.Scrollbar(inspector, orient="vertical", command=self.preview_canvas.yview, style="Clean.Vertical.TScrollbar")
        preview_scroll.grid(row=1, column=1, sticky="ns", pady=(12, 0))
        self.preview_canvas.configure(yscrollcommand=preview_scroll.set)
        self.preview_body = tk.Frame(self.preview_canvas, bg="#eef4f5")
        self.preview_window = self.preview_canvas.create_window((0, 0), window=self.preview_body, anchor="nw")
        self.preview_body.bind(
            "<Configure>",
            lambda _event: self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all")),
        )
        self.preview_canvas.bind("<Configure>", self._sync_preview_width)

        self.template_meta = tk.Label(
            inspector,
            text="Sin plantilla cargada",
            bg=COLORS["panel"],
            fg=COLORS["muted"],
            font=FONT_SMALL,
            justify="left",
            anchor="w",
        )
        self.template_meta.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        editor = self._surface(work, "Editor HTML")
        editor.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        editor.rowconfigure(1, weight=1)
        editor.columnconfigure(0, weight=1)

        self.template_text = tk.Text(
            editor,
            wrap="none",
            bd=0,
            width=54,
            height=24,
            bg="#0f1720",
            fg="#e8f3f2",
            insertbackground="#e8f3f2",
            selectbackground="#265c62",
            font=FONT_MONO,
            padx=14,
            pady=12,
            undo=True,
        )
        self.template_text.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.template_text.bind("<<Modified>>", self._on_template_modified)
        y_scroll = ttk.Scrollbar(editor, orient="vertical", command=self.template_text.yview, style="Clean.Vertical.TScrollbar")
        y_scroll.grid(row=1, column=1, sticky="ns", pady=(12, 0))
        x_scroll = ttk.Scrollbar(editor, orient="horizontal", command=self.template_text.xview, style="Clean.Vertical.TScrollbar")
        x_scroll.grid(row=2, column=0, sticky="ew")
        self.template_text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        editor_actions = tk.Frame(editor, bg=COLORS["panel"])
        editor_actions.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        editor_actions.columnconfigure(3, weight=1)
        self.btn_save_template = self._button(editor_actions, "Guardar cambios", self.save_template, "primary")
        self.btn_save_template.grid(row=0, column=0, padx=(0, 8))
        self._button(editor_actions, "Recargar", self._load_template_editor, "secondary").grid(row=0, column=1, padx=8)
        self._button(editor_actions, "Exportar copia", self.export_template_copy, "secondary").grid(row=0, column=2, padx=8)
        self.template_dirty_label = tk.Label(editor_actions, text="Sin cambios pendientes", bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL)
        self.template_dirty_label.grid(row=0, column=3, sticky="e")

    def _build_config_page(self) -> None:
        page = self._new_page("config")
        page.columnconfigure(0, weight=1)

        smtp = self._surface(page, "Configuracion SMTP", "Datos del remitente y servidor de salida.")
        smtp.grid(row=0, column=0, sticky="ew")
        for c in range(3):
            smtp.columnconfigure(c, weight=1)
        self._field(smtp, "SMTP host", "SMTP_HOST", 1, 0)
        self._field(smtp, "Puerto", "SMTP_PORT", 1, 1, width=12)
        self._field(smtp, "Correo remitente", "SMTP_USER", 1, 2)
        self._field(smtp, "Contrasena SMTP", "SMTP_PASSWORD", 3, 0, show="*")
        self._field(smtp, "Nombre remitente", "FROM_NAME", 3, 1)
        self._field(smtp, "Responder a", "REPLY_TO", 3, 2)

        sending = self._surface(page, "Parametros de campana", "Asunto, volumen y pausas entre correos.")
        sending.grid(row=1, column=0, sticky="ew", pady=(18, 0))
        for c in range(3):
            sending.columnconfigure(c, weight=1)
        self._field(sending, "Asunto", "SUBJECT", 1, 0)
        self._field(sending, "Limite por batch", "DEFAULT_LIMIT", 1, 1, width=12)
        self._field(sending, "Pausa base entre correos (seg)", "SLEEP_SECONDS", 1, 2, width=12)
        self._field(sending, "Pausa aleatoria extra (seg)", "SLEEP_JITTER_SECONDS", 3, 0, width=12)

        bcc = self._surface(page, "BCC aleatorio", "Distribuye copias ocultas en una parte controlada del lote.")
        bcc.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        for c in range(3):
            bcc.columnconfigure(c, weight=1)
        ttk.Checkbutton(bcc, text="Activar BCC aleatorio", variable=self.bcc_enabled, command=self.refresh_counts, style="Modern.TCheckbutton").grid(row=1, column=0, sticky="w", pady=(14, 0))
        self._field(bcc, "Correos BCC", "BCC_RECIPIENTS", 2, 0, columnspan=2)
        self._field(bcc, "BCC por cada 100", "BCC_RANDOM_PER_100", 2, 2, width=12)

        attachments = self._surface(page, "Adjuntos", "Activa o desactiva la inclusión automática de archivos de la carpeta adjuntos.")
        attachments.grid(row=3, column=0, sticky="ew", pady=(18, 0))
        attachments.columnconfigure(0, weight=1)
        ttk.Checkbutton(attachments, text="Incluir adjuntos en pruebas y envíos", variable=self.attachments_enabled, style="Modern.TCheckbutton").grid(row=1, column=0, sticky="w", pady=(14, 0))

        actions = tk.Frame(page, bg=COLORS["app"])
        actions.grid(row=4, column=0, sticky="ew", pady=(18, 0))
        self._button(actions, "Guardar configuracion", self.save_config, "primary").pack(side="left")
        self._button(actions, "Abrir .env", lambda: self.open_path(ROOT / ".env"), "secondary").pack(side="left", padx=10)

    def _build_send_page(self) -> None:
        page = self._new_page("envio")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)

        test = self._surface(page, "Prueba y envio aprobado", "El batch se bloquea hasta que una prueba termine bien y sea aprobada.")
        test.grid(row=0, column=0, sticky="ew")
        for c in range(4):
            test.columnconfigure(c, weight=1)
        self._field(test, "Correo de prueba", "TEST_EMAIL", 1, 0)
        self.btn_test = self._button(test, "1. Enviar prueba", self.send_test, "primary")
        self.btn_test.grid(row=2, column=1, sticky="ew", padx=8, pady=(8, 0))
        self.btn_approve = self._button(test, "2. Aprobar prueba", self.approve_test, "secondary")
        self.btn_approve.grid(row=2, column=2, sticky="ew", padx=8, pady=(8, 0))
        self.btn_batch = self._button(test, "3. Enviar batch", self.send_batch, "secondary")
        self.btn_batch.grid(row=2, column=3, sticky="ew", padx=8, pady=(8, 0))

        split = tk.Frame(page, bg=COLORS["app"])
        split.grid(row=1, column=0, sticky="nsew", pady=(18, 0))
        split.columnconfigure(0, weight=1)
        split.columnconfigure(1, weight=1)
        split.rowconfigure(0, weight=1)

        rules = self._surface(split, "Reglas operativas", "Resumen de seguridad antes de mandar correos reales.")
        rules.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        rules_text = (
            "- La prueba no cambia contactos.csv.\n"
            "- El batch solo usa registros con estado Pendiente.\n"
            "- Al enviar, cada registro queda como Enviado o Error.\n"
            "- La plantilla activa se pasa al motor de envio.\n"
            "- Si el CSV esta abierto en Excel, el motor espera para guardar cambios."
        )
        tk.Label(rules, text=rules_text, bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_UI, justify="left").grid(row=1, column=0, sticky="nw", pady=(14, 0))

        live = self._surface(split, "Estado del envio", "Controles y senales de la sesion actual.")
        live.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.send_status_label = tk.Label(live, text="Listo para enviar prueba", bg=COLORS["panel"], fg=COLORS["ink"], font=("Segoe UI Semibold", 13), justify="left")
        self.send_status_label.grid(row=1, column=0, sticky="w", pady=(14, 4))
        self.send_hint_label = tk.Label(live, text="Guarda configuracion y revisa la plantilla antes de probar.", bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_UI, justify="left", wraplength=420)
        self.send_hint_label.grid(row=2, column=0, sticky="w")

    def _build_contacts_page(self) -> None:
        page = self._new_page("contactos")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(3, weight=1)

        actions = self._surface(page, "Base de contactos", "Lectura rapida del CSV y accesos de mantenimiento.")
        actions.grid(row=0, column=0, sticky="ew")
        self._button(actions, "Actualizar", self.refresh_counts, "primary").grid(row=1, column=0, padx=(0, 8), pady=(14, 0), sticky="w")
        self._button(actions, "Abrir contactos.csv", lambda: self.open_path(ROOT / "contactos.csv"), "secondary").grid(row=1, column=1, padx=8, pady=(14, 0), sticky="w")
        self._button(actions, "Abrir adjuntos", lambda: self.open_path(ROOT / "adjuntos"), "secondary").grid(row=1, column=2, padx=8, pady=(14, 0), sticky="w")

        manager = self._surface(page, "Gestion de estados", "Revisa el contenido del CSV y cambia el estado de filas individuales o de todo el listado.")
        manager.grid(row=1, column=0, sticky="ew", pady=(18, 0))
        manager.columnconfigure(0, weight=1)
        manager.columnconfigure(1, weight=0)
        manager.columnconfigure(2, weight=0)
        self.contacts_status_var = tk.StringVar(value="Pendiente")
        statuses = ["Pendiente", "Enviado", "Error", "Baja", "Cancelado", "Revisar", "Revisar datos"]
        ttk.Combobox(manager, textvariable=self.contacts_status_var, values=statuses, state="readonly", width=20).grid(row=1, column=0, sticky="w", pady=(14, 0))
        self._button(manager, "Aplicar al seleccionado", self.apply_contact_status, "primary").grid(row=1, column=1, padx=(12, 8), pady=(14, 0), sticky="w")
        self._button(manager, "Cambiar todos", self.apply_all_contact_status, "secondary").grid(row=1, column=2, pady=(14, 0), sticky="w")
        tk.Label(manager, text="Selecciona una fila para modificar su estado o usa 'Cambiar todos' para marcar toda la base.", bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL, justify="left", wraplength=720).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))

        summary = self._surface(page, "Estados detectados", "Distribucion actual dentro de contactos.csv.")
        summary.grid(row=2, column=0, sticky="nsew", pady=(18, 0))
        summary.rowconfigure(1, weight=1)
        summary.columnconfigure(0, weight=1)
        self.state_text = tk.Text(summary, height=8, wrap="word", bd=0, bg="#f8fbfb", fg=COLORS["ink"], font=FONT_MONO, padx=14, pady=12)
        self.state_text.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.state_text.configure(state="disabled")

        list_frame = self._surface(page, "Contenido de contactos", "Vista del CSV con columnas clave y estado actual.")
        list_frame.grid(row=3, column=0, sticky="nsew", pady=(18, 0))
        list_frame.rowconfigure(1, weight=1)
        list_frame.columnconfigure(0, weight=1)
        self.contacts_tree = ttk.Treeview(
            list_frame,
            columns=("id", "empresa", "contacto", "destinatarios", "estado"),
            show="headings",
            height=16,
        )
        self.contacts_tree.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        self.contacts_tree.heading("id", text="ID")
        self.contacts_tree.heading("empresa", text="Empresa")
        self.contacts_tree.heading("contacto", text="Contacto")
        self.contacts_tree.heading("destinatarios", text="Destinatarios")
        self.contacts_tree.heading("estado", text="Estado")
        self.contacts_tree.column("id", width=70, stretch=False)
        self.contacts_tree.column("empresa", width=180, stretch=True)
        self.contacts_tree.column("contacto", width=220, stretch=True)
        self.contacts_tree.column("destinatarios", width=260, stretch=True)
        self.contacts_tree.column("estado", width=140, stretch=False)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.contacts_tree.yview)
        scrollbar.grid(row=1, column=1, sticky="ns", pady=(12, 0))
        self.contacts_tree.configure(yscrollcommand=scrollbar.set)
        self.contacts_tree.bind("<<TreeviewSelect>>", self._on_contact_select)
        self.contacts_item_indexes: dict[str, int] = {}
        self.contacts_rows_cache: list[dict[str, str]] = []
        self.refresh_contacts_view()

    def _on_contact_select(self, _event=None) -> None:
        selection = self.contacts_tree.selection()
        if not selection:
            return
        item_id = selection[0]
        index = self.contacts_item_indexes.get(item_id)
        if index is None:
            return
        row = self.contacts_rows_cache[index] if 0 <= index < len(self.contacts_rows_cache) else {}
        current_state = (row.get("estado") or "").strip() or "Pendiente"
        self.contacts_status_var.set(current_state)

    def refresh_contacts_view(self, rows: list[dict[str, str]] | None = None) -> None:
        if rows is None:
            try:
                rows, _ = core.read_contacts(ROOT / "contactos.csv")
            except Exception:
                rows = []
        self.contacts_rows_cache = rows
        self.contacts_item_indexes.clear()
        for child in self.contacts_tree.get_children():
            self.contacts_tree.delete(child)

        for index, row in enumerate(rows):
            values = (
                row.get("id", ""),
                row.get("empresa", ""),
                row.get("contacto", ""),
                row.get("destinatarios", ""),
                row.get("estado", "") or "Pendiente",
            )
            item_id = self.contacts_tree.insert("", "end", values=values)
            self.contacts_item_indexes[item_id] = index

    def apply_contact_status(self) -> None:
        if not self.contacts_rows_cache:
            messagebox.showinfo("Sin contactos", "No hay contactos para actualizar.")
            return
        selection = self.contacts_tree.selection()
        if not selection:
            messagebox.showwarning("Sin selección", "Selecciona una fila para aplicar el estado.")
            return
        state = self.contacts_status_var.get().strip() or "Pendiente"
        item_id = selection[0]
        index = self.contacts_item_indexes.get(item_id)
        if index is None:
            return
        core.update_contact_status(self.contacts_rows_cache, index, state)
        try:
            core.write_contacts(ROOT / "contactos.csv", self.contacts_rows_cache)
            self.refresh_contacts_view(self.contacts_rows_cache)
            self.refresh_counts()
            self.log(f"Estado actualizado para la fila {index + 1}: {state}")
            messagebox.showinfo("Estado actualizado", f"Se actualizó la fila seleccionada a: {state}")
        except Exception as exc:
            messagebox.showerror("No se pudo guardar", str(exc))

    def apply_all_contact_status(self) -> None:
        if not self.contacts_rows_cache:
            messagebox.showinfo("Sin contactos", "No hay contactos para actualizar.")
            return
        state = self.contacts_status_var.get().strip() or "Pendiente"
        try:
            core.bulk_update_contact_status(self.contacts_rows_cache, state)
            core.write_contacts(ROOT / "contactos.csv", self.contacts_rows_cache)
            self.refresh_contacts_view(self.contacts_rows_cache)
            self.refresh_counts()
            self.log(f"Se actualizó el estado de todos los contactos a: {state}")
            messagebox.showinfo("Estado actualizado", f"Se actualizó el estado de todos los contactos a: {state}")
        except Exception as exc:
            messagebox.showerror("No se pudo guardar", str(exc))

    def _build_log_page(self) -> None:
        page = self._new_page("bitacora")
        page.rowconfigure(0, weight=1)
        page.columnconfigure(0, weight=1)

        log_frame = self._surface(page, "Bitacora", "Salida del proceso y eventos de la interfaz.")
        log_frame.grid(row=0, column=0, sticky="nsew")
        log_frame.rowconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            wrap="word",
            bd=0,
            bg=COLORS["terminal"],
            fg="#e8f3f2",
            insertbackground="#e8f3f2",
            selectbackground="#2d5960",
            font=FONT_MONO,
            padx=14,
            pady=12,
        )
        self.log_text.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview, style="Clean.Vertical.TScrollbar")
        scroll.grid(row=1, column=1, sticky="ns", pady=(12, 0))
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log("Listo. Configura SMTP, revisa la plantilla y envia una prueba.")

    def _new_page(self, key: str) -> tk.Frame:
        frame = tk.Frame(self.page_host, bg=COLORS["app"])
        frame.grid(row=0, column=0, sticky="nsew")
        self.pages[key] = frame
        return frame

    def _surface(self, parent: tk.Widget, title: str, subtitle: str | None = None) -> tk.Frame:
        outer = tk.Frame(
            parent,
            bg=COLORS["panel"],
            bd=0,
            relief="flat",
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["line"],
            highlightthickness=1,
        )
        outer.configure(padx=18, pady=16)
        outer.columnconfigure(0, weight=1)
        tk.Label(outer, text=title, bg=COLORS["panel"], fg=COLORS["ink"], font=FONT_SECTION).grid(row=0, column=0, sticky="w")
        if subtitle:
            tk.Label(outer, text=subtitle, bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL).grid(row=0, column=1, sticky="e")
        return outer

    def _metric_card(self, parent: tk.Widget, label: str, key: str, col: int, accent: str) -> None:
        var = tk.StringVar(value="--")
        self.metric_vars[key] = var
        card = tk.Frame(
            parent,
            bg=COLORS["panel"],
            bd=0,
            relief="flat",
            highlightbackground=COLORS["line"],
            highlightcolor=COLORS["line"],
            highlightthickness=1,
            padx=16,
            pady=14,
        )
        card.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 10, 0))
        tk.Label(card, text=label, bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL).pack(anchor="w")
        tk.Label(card, textvariable=var, bg=COLORS["panel"], fg=accent, font=("Segoe UI Semibold", 20)).pack(anchor="w", pady=(6, 0))

    def _button(self, parent: tk.Widget, text: str, command, variant: str = "secondary") -> tk.Button:
        palette = {
            "primary": (COLORS["accent"], "#ffffff", COLORS["accent_dark"]),
            "secondary": ("#edf4f5", COLORS["ink"], "#dfecee"),
            "danger": (COLORS["danger"], "#ffffff", "#842f2f"),
        }
        bg, fg, active = palette.get(variant, palette["secondary"])
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active,
            activeforeground=fg,
            bd=0,
            relief="flat",
            padx=14,
            pady=9,
            cursor="hand2",
            font=("Segoe UI Semibold", 9),
        )

    def _field(
        self,
        parent: tk.Widget,
        label: str,
        key: str,
        row: int,
        col: int,
        width: int = 34,
        show: str | None = None,
        columnspan: int = 1,
    ) -> ttk.Entry:
        tk.Label(parent, text=label, bg=COLORS["panel"], fg=COLORS["ink"], font=("Segoe UI Semibold", 9)).grid(row=row, column=col, sticky="w", padx=8, pady=(14, 4), columnspan=columnspan)
        var = self.vars.get(key) or tk.StringVar()
        self.vars[key] = var
        entry = ttk.Entry(parent, textvariable=var, width=width, show=show or "", style="Modern.TEntry")
        entry.grid(row=row + 1, column=col, sticky="ew", padx=8, pady=(0, 4), columnspan=columnspan)
        return entry

    def show_page(self, key: str) -> None:
        titles = {
            "panel": ("Panel general", "Estado del lote, acciones rapidas y flujo recomendado."),
            "plantilla": ("Centro de plantilla", "Previsualiza, cambia, carga y edita la plantilla HTML activa."),
            "config": ("Configuracion", "Credenciales SMTP, asunto, pausas y BCC."),
            "envio": ("Prueba y envio", "Ejecuta una prueba, apruebala y lanza el batch."),
            "contactos": ("Contactos y adjuntos", "CSV, estados detectados y carpeta de archivos adjuntos."),
            "bitacora": ("Bitacora", "Salida tecnica del proceso."),
        }
        self.current_page = key
        for name, frame in self.pages.items():
            if name == key:
                frame.tkraise()
        for name, btn in self.nav_buttons.items():
            if name == key:
                btn.configure(bg=COLORS["sidebar_soft"], fg="#ffffff")
            else:
                btn.configure(bg=COLORS["sidebar"], fg="#d7e7e8")
        title, subtitle = titles.get(key, ("U4U Mailer Studio", ""))
        self.page_title.configure(text=title)
        self.page_subtitle.configure(text=subtitle)

    def _load_config(self) -> None:
        env = core.load_env()
        defaults = {
            "SMTP_HOST": "smtp.hostinger.com",
            "SMTP_PORT": "465",
            "SMTP_USER": "amorales@u4u.mx",
            "SMTP_PASSWORD": "",
            "FROM_NAME": "Auri Morales - U4U",
            "REPLY_TO": "amorales@u4u.mx",
            "SUBJECT": "Catalogo U4U para depositos dentales",
            "DEFAULT_LIMIT": "100",
            "SLEEP_SECONDS": "5",
            "SLEEP_JITTER_SECONDS": "5",
            "TEST_EMAIL": "emmanuelquintan2020@gmail.com",
            "TEMPLATE_PATH": "plantilla_u4u.html",
            "BCC_RECIPIENTS": "bvilenski@gmail.com, arturogoldsmit@gmail.com",
            "BCC_RANDOM_PER_100": "10",
        }
        for key, default in defaults.items():
            value = env.get(key, default)
            if key == "SMTP_PASSWORD" and value == "PEGA_AQUI_TU_CONTRASENA_SMTP":
                value = ""
            if key in self.vars:
                self.vars[key].set(value)
        self.bcc_enabled.set(core.parse_bool(env.get("BCC_RANDOM_ENABLED", "true"), True))
        self.attachments_enabled.set(core.parse_bool(env.get("ATTACHMENTS_ENABLED", "true"), True))
        self.placeholder_map = core.load_placeholder_map(env)

    def save_config(self) -> None:
        self.save_config_silent()
        self.log("Configuracion guardada en .env")
        self.refresh_counts()
        messagebox.showinfo("Configuracion guardada", "Listo. La configuracion quedo guardada en .env.")

    def save_config_silent(self) -> None:
        values = {key: var.get().strip() for key, var in self.vars.items()}
        values["BCC_RANDOM_ENABLED"] = "true" if self.bcc_enabled.get() else "false"
        values["ATTACHMENTS_ENABLED"] = "true" if self.attachments_enabled.get() else "false"
        values["PLACEHOLDER_MAP"] = core.dump_placeholder_map(self.current_placeholder_mapping())
        core.save_env(values)

    def refresh_counts(self) -> None:
        try:
            rows, _ = core.read_contacts(ROOT / "contactos.csv")
            pending = core.remaining_pending(rows)
            state_counts = core.count_by_state(rows)
            try:
                limit = int(self.vars.get("DEFAULT_LIMIT", tk.StringVar(value="100")).get() or "100")
            except ValueError:
                limit = 100
            batch = min(pending, max(1, limit)) if pending else 0
            try:
                per100 = int(self.vars.get("BCC_RANDOM_PER_100", tk.StringVar(value="10")).get() or "10")
            except ValueError:
                per100 = 10
            bcc_count = round(batch * per100 / 100) if self.bcc_enabled.get() else 0

            self.metric_vars["pending"].set(str(pending))
            self.metric_vars["batch"].set(str(batch))
            self.metric_vars["bcc"].set(f"{bcc_count} de {batch}")
            self.metric_vars["template"].set("OK" if self.get_template_path().exists() else "Falta")
            self.status_pill.configure(text="Listo", bg="#e7f4f2", fg=COLORS["accent_dark"])

            lines = [f"{state:<18} {count:>5}" for state, count in sorted(state_counts.items())]
            self._replace_text(self.state_text, "\n".join(lines) if lines else "No hay contactos detectados.")
            self.refresh_contacts_view(rows)
        except Exception as exc:
            self.metric_vars.get("pending", tk.StringVar()).set("Error")
            self.metric_vars.get("batch", tk.StringVar()).set("--")
            self.metric_vars.get("bcc", tk.StringVar()).set("--")
            self.log(f"Error al leer contactos.csv: {exc}")
        self._update_template_summary()

    def log(self, message: str) -> None:
        line = message.rstrip() + "\n"
        for widget_name in ("log_text", "dashboard_log"):
            widget = getattr(self, widget_name, None)
            if not widget:
                continue
            state = str(widget.cget("state"))
            if state == "disabled":
                widget.configure(state="normal")
            widget.insert("end", line)
            widget.see("end")
            if widget_name == "dashboard_log":
                content = widget.get("1.0", "end-1c").splitlines()
                if len(content) > 80:
                    widget.delete("1.0", f"{len(content) - 80}.0")
                widget.configure(state="disabled")

    def _replace_text(self, widget: tk.Text, value: str) -> None:
        state = str(widget.cget("state"))
        if state == "disabled":
            widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", value)
        if state == "disabled":
            widget.configure(state="disabled")

    def get_csv_columns(self) -> list[str]:
        try:
            rows, _ = core.read_contacts(ROOT / "contactos.csv")
            columns: list[str] = []
            for row in rows:
                for column in row.keys():
                    if column and column not in columns:
                        columns.append(column)
                if columns:
                    break
            return columns or list(core.REQUIRED_COLUMNS)
        except Exception:
            return list(core.REQUIRED_COLUMNS)

    def current_placeholder_mapping(self) -> dict[str, str]:
        mapping = dict(self.placeholder_map)
        for placeholder, var in self.placeholder_vars.items():
            value = var.get().strip()
            if value:
                mapping[placeholder] = value
        return mapping

    def current_env_values(self) -> dict[str, str]:
        values = {key: var.get().strip() for key, var in self.vars.items()}
        values["BCC_RANDOM_ENABLED"] = "true" if self.bcc_enabled.get() else "false"
        values["ATTACHMENTS_ENABLED"] = "true" if self.attachments_enabled.get() else "false"
        values["PLACEHOLDER_MAP"] = core.dump_placeholder_map(self.current_placeholder_mapping())
        return values

    def preview_sample_row(self) -> dict[str, str]:
        row = {column: "" for column in self.csv_columns}
        try:
            rows, _ = core.read_contacts(ROOT / "contactos.csv")
            if rows:
                row.update(rows[0])
        except Exception:
            pass
        row["contacto"] = "JOSE EMMANUEL QUINTANA TORRES"
        row["destinatarios"] = self.vars.get("TEST_EMAIL", tk.StringVar(value="")).get().strip()
        row.setdefault("empresa", "U4U")
        return row

    def _preview_replaced_html(self, html: str) -> str:
        return core.personalize_template(html, self.preview_sample_row(), self.current_env_values())

    def _sample_value_for_column(self, column: str) -> str:
        value = core.value_for_placeholder(self.preview_sample_row(), column)
        return value or "sin dato"

    def refresh_placeholder_controls(self, placeholders: list[str]) -> None:
        if not hasattr(self, "placeholder_map_frame"):
            return
        self.csv_columns = self.get_csv_columns()
        self.placeholder_map = self.current_placeholder_mapping()
        for child in self.placeholder_map_frame.winfo_children():
            child.destroy()
        self.placeholder_vars.clear()
        self.placeholder_rows.clear()

        if not placeholders:
            self.placeholder_status.configure(text="No detecte placeholders tipo [Campo] en esta plantilla.")
            return

        saved_map = core.load_placeholder_map(
            {"PLACEHOLDER_MAP": core.dump_placeholder_map(self.placeholder_map)},
            placeholders,
            self.csv_columns,
        )
        self.placeholder_status.configure(text=f"{len(placeholders)} placeholder(s) detectado(s). Cada uno se reemplazara con la columna seleccionada.")
        tk.Label(self.placeholder_map_frame, text="Placeholder", bg=COLORS["panel"], fg=COLORS["muted"], font=("Segoe UI Semibold", 9)).grid(row=0, column=0, sticky="w", padx=(0, 12), pady=(0, 6))
        tk.Label(self.placeholder_map_frame, text="Columna CSV", bg=COLORS["panel"], fg=COLORS["muted"], font=("Segoe UI Semibold", 9)).grid(row=0, column=1, sticky="w", padx=12, pady=(0, 6))
        tk.Label(self.placeholder_map_frame, text="Valor en prueba", bg=COLORS["panel"], fg=COLORS["muted"], font=("Segoe UI Semibold", 9)).grid(row=0, column=2, sticky="w", padx=(12, 0), pady=(0, 6))

        for row_index, placeholder in enumerate(placeholders, start=1):
            selected = saved_map.get(placeholder) or core.default_column_for_placeholder(placeholder, self.csv_columns)
            var = tk.StringVar(value=selected)
            self.placeholder_vars[placeholder] = var
            tk.Label(self.placeholder_map_frame, text=placeholder, bg=COLORS["panel"], fg=COLORS["ink"], font=FONT_SMALL).grid(row=row_index, column=0, sticky="w", padx=(0, 12), pady=3)
            combo = ttk.Combobox(self.placeholder_map_frame, textvariable=var, values=self.csv_columns, state="readonly", width=22)
            combo.grid(row=row_index, column=1, sticky="ew", padx=12, pady=3)
            sample_var = tk.StringVar(value=self._sample_value_for_column(selected))
            sample_label = tk.Label(self.placeholder_map_frame, textvariable=sample_var, bg=COLORS["panel"], fg=COLORS["muted"], font=FONT_SMALL, anchor="w")
            sample_label.grid(row=row_index, column=2, sticky="ew", padx=(12, 0), pady=3)

            def changed(_event=None, key=placeholder, value_var=var, sample=sample_var) -> None:
                self.placeholder_map[key] = value_var.get().strip()
                sample.set(self._sample_value_for_column(value_var.get().strip()))
                self.save_config_silent()
                self._update_template_inspector(self.template_text.get("1.0", "end-1c"), refresh_mapping=False)

            combo.bind("<<ComboboxSelected>>", changed)

        self.placeholder_map = self.current_placeholder_mapping()

    def save_placeholder_mapping(self) -> None:
        self.placeholder_map = self.current_placeholder_mapping()
        self.save_config_silent()
        self._update_template_inspector(self.template_text.get("1.0", "end-1c"), refresh_mapping=False)
        messagebox.showinfo("Mapeo guardado", "El mapeo de placeholders quedo guardado en .env.")

    def get_template_path(self) -> Path:
        raw = self.vars.get("TEMPLATE_PATH", tk.StringVar(value="plantilla_u4u.html")).get().strip()
        if not raw:
            return DEFAULT_TEMPLATE
        path = Path(raw)
        if not path.is_absolute():
            path = ROOT / path
        return path

    def _display_template_path(self) -> str:
        path = self.get_template_path()
        try:
            return str(path.relative_to(ROOT))
        except ValueError:
            return str(path)

    def _load_template_editor(self) -> None:
        path = self.get_template_path()
        try:
            html = path.read_text(encoding="utf-8")
        except FileNotFoundError:
            html = ""
            self.log(f"No encontre la plantilla activa: {path}")
        except Exception as exc:
            html = ""
            self.log(f"Error al cargar plantilla: {exc}")
        self.template_text.configure(state="normal")
        self.template_text.delete("1.0", "end")
        self.template_text.insert("1.0", html)
        self.template_text.yview_moveto(0)
        self.template_text.xview_moveto(0)
        self.template_text.edit_modified(False)
        self.template_dirty = False
        self._update_template_dirty_state()
        self._update_template_inspector(html)
        self._update_template_summary()

    def _on_template_modified(self, _event=None) -> None:
        if self.template_text.edit_modified():
            self.template_dirty = True
            self._update_template_dirty_state()
            html = self.template_text.get("1.0", "end-1c")
            self._update_template_inspector(html)
            self.template_text.edit_modified(False)

    def _update_template_dirty_state(self) -> None:
        if self.template_dirty:
            self.template_dirty_label.configure(text="Cambios sin guardar", fg=COLORS["warning"])
        else:
            self.template_dirty_label.configure(text="Sin cambios pendientes", fg=COLORS["muted"])

    def _update_template_summary(self) -> None:
        path = self.get_template_path()
        self.sidebar_template.configure(text=self._display_template_path())
        if "template" in self.metric_vars:
            self.metric_vars["template"].set("OK" if path.exists() else "Falta")

    def _sync_preview_width(self, event=None) -> None:
        if not hasattr(self, "preview_canvas") or not hasattr(self, "preview_window"):
            return
        width = max(360, event.width if event is not None else self.preview_canvas.winfo_width())
        self.preview_canvas.itemconfigure(self.preview_window, width=width)
        self.preview_canvas.xview_moveto(0)
        if hasattr(self, "template_meta"):
            self.template_meta.configure(wraplength=max(320, width - 24))

    def _rerender_preview_after_resize(self) -> None:
        self._preview_resize_job = None
        if self.preview_html:
            self._schedule_browser_preview(delay=120, reset_scroll=False)
            self.after_idle(self._sync_preview_width)

    def _show_preview_message(self, title: str, detail: str = "") -> None:
        for child in self.preview_body.winfo_children():
            child.destroy()
        self.preview_image = None
        box = tk.Frame(self.preview_body, bg="#eef4f5", padx=24, pady=24)
        box.pack(fill="both", expand=True)
        tk.Label(box, text=title, bg="#eef4f5", fg=COLORS["ink"], font=("Segoe UI Semibold", 15)).pack(anchor="center", pady=(80, 6))
        if detail:
            tk.Label(box, text=detail, bg="#eef4f5", fg=COLORS["muted"], font=FONT_UI, wraplength=420, justify="center").pack(anchor="center")
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))

    def _find_browser_executables(self) -> list[str]:
        candidates = [
            shutil.which("chrome"),
            shutil.which("chromium"),
            os.path.join(os.environ.get("ProgramFiles", ""), "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Google", "Chrome", "Application", "chrome.exe"),
            shutil.which("msedge"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft", "Edge", "Application", "msedge.exe"),
        ]
        browsers: list[str] = []
        for candidate in candidates:
            if candidate and Path(candidate).exists() and candidate not in browsers:
                browsers.append(candidate)
        return browsers

    def _find_browser_executable(self) -> str | None:
        browsers = self._find_browser_executables()
        return browsers[0] if browsers else None

    def _write_preview_html(self, html: str) -> Path:
        preview_path = ROOT / "logs" / "preview_plantilla.html"
        preview_path.parent.mkdir(exist_ok=True)
        preview_path.write_text(html, encoding="utf-8")
        return preview_path

    def _schedule_browser_preview(self, delay: int = 350, reset_scroll: bool = True) -> None:
        if self._preview_render_job is not None:
            self.after_cancel(self._preview_render_job)
        self._preview_render_job = self.after(delay, lambda: self._start_browser_preview(reset_scroll=reset_scroll))

    def _start_browser_preview(self, reset_scroll: bool = True) -> None:
        self._preview_render_job = None
        html = self.preview_html
        if not html.strip():
            self._show_preview_message("Sin plantilla cargada", "Carga una plantilla HTML para verla aqui renderizada.")
            return
        browsers = self._find_browser_executables()
        if not browsers:
            self._show_preview_message("No encontre navegador", "Instala Microsoft Edge o Chrome para renderizar la vista previa real.")
            return

        self._preview_request_id += 1
        request_id = self._preview_request_id
        canvas_width = max(420, self.preview_canvas.winfo_width() - 10)
        viewport_width = max(420, min(920, canvas_width))
        viewport_height = 1800
        self._preview_render_width = viewport_width
        html_path = self._write_preview_html(html)
        png_path = ROOT / "logs" / f"preview_plantilla_{request_id}.png"
        self._show_preview_message("Renderizando plantilla", "Generando captura con el motor del navegador.")
        self.update_idletasks()

        def worker() -> None:
            errors = []
            try:
                for browser in browsers:
                    try:
                        self._capture_html_with_browser(browser, html_path, png_path, viewport_width, viewport_height)
                        break
                    except Exception as exc:
                        errors.append(f"{Path(browser).name}: {exc}")
                else:
                    raise RuntimeError(" | ".join(errors))
                self.queue.put(("preview_ready", (request_id, str(png_path), reset_scroll)))
            except Exception as exc:
                self.queue.put(("preview_error", (request_id, str(exc))))

        threading.Thread(target=worker, daemon=True).start()

    def _capture_html_with_browser(self, browser: str, html_path: Path, png_path: Path, width: int, height: int) -> None:
        png_path.unlink(missing_ok=True)
        creationflags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
        fallback_screenshot = png_path.parent / "screenshot.png"
        fallback_screenshot.unlink(missing_ok=True)
        profile_dir = tempfile.mkdtemp(prefix="u4u_preview_")
        try:
            base_args = [
                browser,
                "--disable-gpu",
                "--disable-extensions",
                "--disable-background-networking",
                "--hide-scrollbars",
                "--no-first-run",
                "--no-default-browser-check",
                f"--user-data-dir={profile_dir}",
                "--force-device-scale-factor=1",
                "--run-all-compositor-stages-before-draw",
                "--virtual-time-budget=3000",
                f"--window-size={width},{height}",
            ]
            attempts = [
                base_args + ["--headless=new", f"--screenshot={png_path}", html_path.as_uri()],
                base_args + ["--headless", f"--screenshot={png_path}", html_path.as_uri()],
                base_args + ["--headless=new", "--screenshot", html_path.as_uri()],
                base_args + ["--headless", "--screenshot", html_path.as_uri()],
            ]
            last_error = ""
            for args in attempts:
                proc = subprocess.run(
                    args,
                    cwd=str(png_path.parent),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=30,
                    creationflags=creationflags,
                )
                for candidate in (png_path, fallback_screenshot):
                    if candidate.exists() and candidate.stat().st_size > 0:
                        if candidate != png_path:
                            shutil.move(str(candidate), str(png_path))
                        return
                output = (proc.stderr or proc.stdout or "").strip()
                last_error = output or f"Chrome/Edge termino con codigo {proc.returncode}, pero no genero la captura."
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)
        raise RuntimeError(last_error or "No se pudo generar la captura de la plantilla.")

    def _display_preview_image(self, png_path: str, reset_scroll: bool) -> None:
        for child in self.preview_body.winfo_children():
            child.destroy()
        self.preview_image = tk.PhotoImage(file=png_path)
        holder = tk.Frame(self.preview_body, bg="#eef4f5", padx=0, pady=0)
        holder.pack(fill="both", expand=True)
        tk.Label(holder, image=self.preview_image, bg="#eef4f5", bd=0).pack(anchor="n")
        self.preview_canvas.update_idletasks()
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox("all"))
        self.preview_canvas.xview_moveto(0)
        if reset_scroll:
            self.preview_canvas.yview_moveto(0)

    def _clean_html_fragment(self, fragment: str) -> str:
        text = re.sub(r"<br\s*/?>", "\n", fragment, flags=re.I)
        text = re.sub(r"</p>|</h[1-6]>|</td>", "\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)
        text = html_lib.unescape(text)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n\s+", "\n", text)
        return re.sub(r"\n{2,}", "\n", text).strip()

    def _extract_template_preview_data(self, html: str) -> dict[str, object]:
        sample = {
            "empresa": "Deposito Dental Reforma",
            "contacto": "Mariana Torres",
            "destinatarios": self.vars.get("TEST_EMAIL", tk.StringVar()).get(),
            "estado": "Pendiente",
        }
        personalized = core.personalize_template(html, sample)
        title_match = re.search(r"<title[^>]*>(.*?)</title>", personalized, flags=re.I | re.S)
        h1_match = re.search(r"<h1[^>]*>(.*?)</h1>", personalized, flags=re.I | re.S)
        img_alts = [
            self._clean_html_fragment(match)
            for match in re.findall(r"<img[^>]*\salt=[\"']([^\"']+)[\"'][^>]*>", personalized, flags=re.I | re.S)
        ]
        paragraphs = [
            self._clean_html_fragment(match)
            for match in re.findall(r"<p[^>]*>(.*?)</p>", personalized, flags=re.I | re.S)
        ]
        paragraphs = [item for item in paragraphs if item]
        buttons = [
            self._clean_html_fragment(match)
            for match in re.findall(r"<a[^>]*>(.*?)</a>", personalized, flags=re.I | re.S)
        ]
        buttons = [item for item in buttons if item and len(item) <= 42]
        plain_lines = [line.strip() for line in core.html_to_plain_text(personalized).splitlines() if line.strip()]

        greeting = next((item for item in paragraphs if item.lower().startswith("hola") or item.startswith("¡Hola")), "")
        footer_terms = ("recibiste", "no deseas", "responde", "catalogo pdf", "catálogo pdf", "saludos")
        body = [
            item
            for item in paragraphs
            if item != greeting and not any(term in item.lower() for term in footer_terms)
        ][:4]
        footer = [line for line in plain_lines if line.startswith("Saludos") or line in {"Auri Morales", "U4U Uniformes"}][:3]

        return {
            "title": self._clean_html_fragment(title_match.group(1)) if title_match else "Plantilla U4U",
            "headline": self._clean_html_fragment(h1_match.group(1)) if h1_match else "U4U",
            "hero_alt": next((alt for alt in img_alts if "catalog" in alt.lower() or "u4u" in alt.lower()), ""),
            "greeting": greeting or "Hola Mariana Torres",
            "body": body,
            "buttons": buttons[:2] or ["Solicitar informacion", "Visitar u4u.mx"],
            "footer": footer,
            "accent": "#2a9f9f" if "#2a9f9f" in personalized.lower() else COLORS["accent"],
        }

    def _render_email_preview(self, html: str, reset_scroll: bool = True) -> None:
        for child in self.preview_body.winfo_children():
            child.destroy()
        if not html.strip():
            empty = tk.Frame(self.preview_body, bg="#eef4f5", padx=24, pady=24)
            empty.pack(fill="both", expand=True)
            tk.Label(empty, text="Sin plantilla cargada", bg="#eef4f5", fg=COLORS["ink"], font=("Segoe UI Semibold", 16)).pack(anchor="center", pady=(80, 6))
            tk.Label(empty, text="Carga una plantilla HTML para ver aqui una simulacion visual.", bg="#eef4f5", fg=COLORS["muted"], font=FONT_UI).pack(anchor="center")
            return

        data = self._extract_template_preview_data(html)
        canvas_width = max(420, self.preview_canvas.winfo_width())
        self._preview_render_width = canvas_width
        mail_width = max(360, min(620, canvas_width - 34))
        wrap = max(270, mail_width - 72)
        outer_pad = max(8, (canvas_width - mail_width) // 2)
        outer = tk.Frame(self.preview_body, bg="#eef4f5", padx=outer_pad, pady=14)
        outer.pack(fill="both", expand=True)

        email = tk.Frame(
            outer,
            bg="#ffffff",
            highlightbackground="#dce8ea",
            highlightcolor="#dce8ea",
            highlightthickness=1,
        )
        email.pack(fill="x", anchor="n")

        header = tk.Frame(email, bg=str(data["accent"]), padx=24, pady=18)
        header.pack(fill="x")
        logo = tk.Label(header, text="U4U", bg="#ffffff", fg=str(data["accent"]), font=("Segoe UI Semibold", 16), width=6, pady=9)
        logo.pack(anchor="center", pady=(0, 10))
        tk.Label(
            header,
            text=str(data["headline"]),
            bg=str(data["accent"]),
            fg="#ffffff",
            font=("Segoe UI Semibold", 14),
            wraplength=wrap,
            justify="center",
        ).pack(anchor="center", fill="x")

        hero = tk.Frame(email, bg="#f2fbfb", padx=24, pady=22)
        hero.pack(fill="x")
        tk.Label(
            hero,
            text="Te compartimos\nnuestro catalogo",
            bg="#f2fbfb",
            fg="#07354b",
            font=("Segoe UI Semibold", 22),
            justify="left",
            wraplength=wrap,
            anchor="w",
        ).pack(anchor="w", fill="x")
        hero_alt = str(data["hero_alt"])
        if hero_alt:
            tk.Label(hero, text=hero_alt, bg="#f2fbfb", fg=COLORS["muted"], font=FONT_SMALL, wraplength=wrap, justify="left", anchor="w").pack(anchor="w", fill="x", pady=(10, 0))
        tk.Label(hero, text="Imagen hero detectada en la plantilla", bg="#dff1f1", fg=COLORS["accent_dark"], font=("Segoe UI Semibold", 9), padx=10, pady=6).pack(anchor="w", pady=(12, 0))

        body = tk.Frame(email, bg="#ffffff", padx=28, pady=24)
        body.pack(fill="x")
        tk.Label(
            body,
            text=str(data["greeting"]),
            bg="#ffffff",
            fg="#0b3653",
            font=("Segoe UI Semibold", 12),
            wraplength=wrap,
            justify="left",
            anchor="w",
        ).pack(anchor="w", fill="x", pady=(0, 12))
        for paragraph in data["body"]:  # type: ignore[index]
            tk.Label(
                body,
                text=str(paragraph),
                bg="#ffffff",
                fg="#263238",
                font=("Segoe UI", 11),
                wraplength=wrap,
                justify="left",
                anchor="w",
            ).pack(anchor="w", fill="x", pady=(0, 12))

        actions = tk.Frame(body, bg="#ffffff")
        actions.pack(anchor="center", pady=(8, 12))
        for index, button_text in enumerate(data["buttons"]):  # type: ignore[index]
            color = str(data["accent"]) if index == 0 else "#0b3653"
            tk.Label(
                actions,
                text=str(button_text),
                bg=color,
                fg="#ffffff",
                font=("Segoe UI Semibold", 10),
                padx=16,
                pady=9,
            ).pack(side="left", padx=5)

        if data["footer"]:
            footer = tk.Frame(email, bg="#f8fbfb", padx=28, pady=18)
            footer.pack(fill="x")
            tk.Label(
                footer,
                text="\n".join(str(item) for item in data["footer"]),  # type: ignore[index]
                bg="#f8fbfb",
                fg="#0b3653",
                font=FONT_SMALL,
                justify="left",
                wraplength=wrap,
                anchor="w",
            ).pack(anchor="w", fill="x")
        if reset_scroll:
            self.preview_canvas.after_idle(lambda: (self.preview_canvas.xview_moveto(0), self.preview_canvas.yview_moveto(0)))
        else:
            self.preview_canvas.after_idle(lambda: self.preview_canvas.xview_moveto(0))

    def _update_template_inspector(self, html: str, refresh_mapping: bool = True) -> None:
        path = self.get_template_path()
        placeholders = core.detect_placeholders(html)
        if refresh_mapping:
            self.refresh_placeholder_controls(placeholders)
        self.preview_html = self._preview_replaced_html(html)
        links = len(re.findall(r"<a\s", html, flags=re.I))
        images = len(re.findall(r"<img\s", html, flags=re.I))
        title_match = re.search(r"<title[^>]*>(.*?)</title>", html, flags=re.I | re.S)
        title = re.sub(r"\s+", " ", title_match.group(1)).strip() if title_match else "Sin title"
        size = len(html.encode("utf-8"))
        modified = "No existe"
        if path.exists():
            modified = path.stat().st_mtime
            import datetime as _dt
            modified = _dt.datetime.fromtimestamp(modified).strftime("%Y-%m-%d %H:%M")
        meta = (
            f"{self._display_template_path()}  |  {title}  |  {size:,} bytes  |  "
            f"Modificado: {modified}  |  Imagenes: {images}  |  Links: {links}  |  "
            f"Placeholders: {', '.join(placeholders) if placeholders else 'ninguno'}"
        )
        self.template_meta.configure(text=meta)
        self._schedule_browser_preview(delay=450, reset_scroll=True)

    def choose_template_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Selecciona una plantilla HTML",
            filetypes=[("HTML", "*.html *.htm"), ("Todos los archivos", "*.*")],
            initialdir=str(ROOT),
        )
        if not file_path:
            return
        self.vars["TEMPLATE_PATH"].set(file_path)
        self.save_config_silent()
        self._load_template_editor()
        self.refresh_counts()
        self.log(f"Plantilla activa cambiada a: {file_path}")

    def import_template_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Importar plantilla a esta carpeta",
            filetypes=[("HTML", "*.html *.htm"), ("Todos los archivos", "*.*")],
            initialdir=str(ROOT),
        )
        if not file_path:
            return
        source = Path(file_path)
        if not source.exists():
            messagebox.showerror("Plantilla no encontrada", "No se encontro el archivo seleccionado.")
            return
        if DEFAULT_TEMPLATE.exists():
            backup = ROOT / f"plantilla_u4u.backup.{DEFAULT_TEMPLATE.stat().st_mtime_ns}.html"
            shutil.copy2(DEFAULT_TEMPLATE, backup)
            self.log(f"Backup creado: {backup.name}")
        shutil.copy2(source, DEFAULT_TEMPLATE)
        self.vars["TEMPLATE_PATH"].set("plantilla_u4u.html")
        self.save_config_silent()
        self._load_template_editor()
        self.refresh_counts()
        self.log(f"Plantilla importada como {DEFAULT_TEMPLATE.name}")

    def save_template(self, silent: bool = False) -> bool:
        path = self.get_template_path()
        html = self.template_text.get("1.0", "end-1c")
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(html, encoding="utf-8")
            self.template_dirty = False
            self.template_text.edit_modified(False)
            self._update_template_dirty_state()
            self._update_template_inspector(html)
            self.refresh_counts()
            self.save_config_silent()
            self.log(f"Plantilla guardada: {path}")
            if not silent:
                messagebox.showinfo("Plantilla guardada", "La plantilla activa quedo guardada.")
            return True
        except Exception as exc:
            self.log(f"Error al guardar plantilla: {exc}")
            if not silent:
                messagebox.showerror("No se pudo guardar", str(exc))
            return False

    def export_template_copy(self) -> None:
        target = filedialog.asksaveasfilename(
            title="Exportar copia de plantilla",
            defaultextension=".html",
            filetypes=[("HTML", "*.html"), ("Todos los archivos", "*.*")],
            initialfile="plantilla_u4u_copia.html",
            initialdir=str(ROOT),
        )
        if not target:
            return
        Path(target).write_text(self.template_text.get("1.0", "end-1c"), encoding="utf-8")
        self.log(f"Copia exportada: {target}")
        messagebox.showinfo("Copia exportada", "La copia de la plantilla quedo guardada.")

    def open_template_preview(self) -> None:
        html = self.template_text.get("1.0", "end-1c")
        if not html.strip():
            path = self.get_template_path()
            if path.exists():
                html = path.read_text(encoding="utf-8")
        if not html.strip():
            messagebox.showwarning("Sin plantilla", "No hay HTML de plantilla para previsualizar.")
            return
        preview_path = self._write_preview_html(self._preview_replaced_html(html))
        self.open_path(preview_path)
        self.log(f"Vista previa abierta: {preview_path}")

    def _confirm_template_saved_for_send(self) -> bool:
        path = self.get_template_path()
        if not path.exists() and not self.template_text.get("1.0", "end-1c").strip():
            messagebox.showerror("Plantilla no encontrada", "La plantilla activa no existe o esta vacia.")
            self.show_page("plantilla")
            return False
        if self.template_dirty:
            save_now = messagebox.askyesno("Cambios sin guardar", "La plantilla tiene cambios sin guardar. Guardarlos antes de enviar?")
            if save_now and not self.save_template(silent=True):
                return False
        return True

    def _set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        self.btn_test.configure(state=state)
        if running:
            self.btn_approve.configure(state="disabled")
            self.btn_batch.configure(state="disabled")
            self.status_pill.configure(text="Ejecutando", bg="#fff3df", fg=COLORS["warning"])
            self.send_status_label.configure(text="Proceso en ejecucion")
            self.send_hint_label.configure(text="Puedes seguir la salida completa en Bitacora.")
        else:
            self.btn_approve.configure(state="normal" if self.test_sent_ok else "disabled")
            self.btn_batch.configure(state="normal" if self.test_approved else "disabled")
            self.status_pill.configure(text="Listo", bg="#e7f4f2", fg=COLORS["accent_dark"])
            if self.test_approved:
                self.send_status_label.configure(text="Prueba aprobada")
                self.send_hint_label.configure(text="El batch esta habilitado. Revisa el limite y confirma el envio.")
            elif self.test_sent_ok:
                self.send_status_label.configure(text="Prueba enviada")
                self.send_hint_label.configure(text="Revisa el correo de prueba y apruebalo si se ve correcto.")
            else:
                self.send_status_label.configure(text="Listo para enviar prueba")
                self.send_hint_label.configure(text="Guarda configuracion y revisa la plantilla antes de probar.")

    def _run_command(self, args: list[str], on_success: str | None = None) -> None:
        self.save_config_silent()
        self._set_running(True)
        self.show_page("bitacora")
        self.log("Ejecutando: " + " ".join(args))

        def worker() -> None:
            try:
                env = os.environ.copy()
                env["PYTHONIOENCODING"] = "utf-8"
                proc = subprocess.Popen(
                    args,
                    cwd=str(ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    env=env,
                )
                assert proc.stdout is not None
                for line in proc.stdout:
                    self.queue.put(("log", line.rstrip()))
                code = proc.wait()
                self.queue.put(("done", (code, on_success)))
            except Exception as exc:
                self.queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def send_test(self) -> None:
        if not self._confirm_template_saved_for_send():
            return
        email = self.vars["TEST_EMAIL"].get().strip() or "emmanuelquintan2020@gmail.com"
        if not core.parse_recipients(email):
            messagebox.showerror("Correo no valido", "El correo de prueba no parece valido.")
            return
        self.test_sent_ok = False
        self.test_approved = False
        args = [PYTHON, "enviar_correos.py", "--template", str(self.get_template_path()), "--test-to", email, "--yes"]
        self._run_command(args, on_success="test")

    def approve_test(self) -> None:
        if not self.test_sent_ok:
            messagebox.showwarning("Primero envia la prueba", "Primero debe enviarse correctamente el correo de prueba.")
            return
        approved = messagebox.askyesno(
            "Aprobar prueba",
            "Ya revisaste el correo de prueba y confirmas que se ve bien?\n\nSi apruebas, se habilita el envio del batch.",
        )
        if approved:
            self.test_approved = True
            self.btn_batch.configure(state="normal")
            self.send_status_label.configure(text="Prueba aprobada")
            self.send_hint_label.configure(text="Ya puedes enviar el batch con la plantilla activa.")
            self.log("Prueba aprobada. Ya puedes enviar el batch.")

    def send_batch(self) -> None:
        if not self.test_approved:
            messagebox.showwarning("Prueba no aprobada", "Primero envia y aprueba el correo de prueba.")
            return
        if not self._confirm_template_saved_for_send():
            return
        self.refresh_counts()
        try:
            limit = int(self.vars["DEFAULT_LIMIT"].get() or "100")
        except ValueError:
            limit = 100
        bcc_text = "con BCC aleatorio" if self.bcc_enabled.get() else "sin BCC aleatorio"
        confirm = messagebox.askyesno(
            "Confirmar envio",
            f"Se enviara el siguiente lote de hasta {limit} registros pendientes, {bcc_text}.\n\nConfirmas enviar ahora?",
        )
        if not confirm:
            self.log("Batch cancelado por el usuario.")
            return
        args = [PYTHON, "enviar_correos.py", "--template", str(self.get_template_path()), "--limit", str(limit), "--yes"]
        if self.bcc_enabled.get():
            args.append("--enable-random-bcc")
        self._run_command(args, on_success="batch")

    def _process_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self.log(str(payload))
                elif kind == "error":
                    self.log("ERROR: " + str(payload))
                    self._set_running(False)
                    messagebox.showerror("Error", str(payload))
                elif kind == "preview_ready":
                    request_id, png_path, reset_scroll = payload  # type: ignore[misc]
                    if int(request_id) == self._preview_request_id:
                        self._display_preview_image(str(png_path), bool(reset_scroll))
                elif kind == "preview_error":
                    request_id, error = payload  # type: ignore[misc]
                    if int(request_id) == self._preview_request_id:
                        self._show_preview_message("No se pudo renderizar", str(error))
                        self.log(f"Error al renderizar vista previa: {error}")
                elif kind == "done":
                    code, on_success = payload  # type: ignore[misc]
                    self.log(f"Proceso terminado con codigo {code}.")
                    if int(code) == 0:
                        if on_success == "test":
                            self.test_sent_ok = True
                            self.test_approved = False
                            self.log("Prueba enviada. Revisa tu correo y pulsa 'Aprobar prueba'.")
                            messagebox.showinfo("Prueba enviada", "Prueba enviada correctamente. Revisa el correo y despues aprueba la prueba.")
                        elif on_success == "batch":
                            self.test_approved = False
                            messagebox.showinfo("Batch terminado", "El batch termino correctamente.")
                    else:
                        messagebox.showerror("Proceso con error", f"El proceso termino con codigo {code}. Revisa la bitacora.")
                    self.refresh_counts()
                    self._set_running(False)
        except queue.Empty:
            pass
        self.after(150, self._process_queue)

    def open_path(self, path: Path) -> None:
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(path))  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("No se pudo abrir", str(exc))


if __name__ == "__main__":
    app = U4UMailerApp()
    app.mainloop()
