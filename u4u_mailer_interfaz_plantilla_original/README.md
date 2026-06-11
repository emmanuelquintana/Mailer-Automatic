# U4U Mailer Studio

╔══════════════════════════════════════════════╗
║  U4U Mailer Studio                          ║
║  Envíos inteligentes, seguros y con estilo ║
╚══════════════════════════════════════════════╝

✨ Un gestor visual para preparar, revisar y enviar correos masivos con plantilla HTML, adjuntos, estados de contacto y control de lotes.

## 🚀 Características

- Interfaz gráfica para administrar SMTP, plantilla y configuración.
- Envío de pruebas antes del lote real.
- Control de estados de contactos: Pendiente, Enviado, Error, Baja, etc.
- Soporte para adjuntos opcionales.
- Vista rápida del contenido del CSV y resumen de estados.

## 🛠️ Requisitos

- Python 3.10 o superior
- Paquetes listados en requirements.txt

## ▶️ Instalación y uso

1. Instala dependencias:
   ```bash
   pip install -r requirements.txt
   ```

2. Crea tu archivo de entorno local:
   ```bash
   copy .env.example .env
   ```

3. Edita .env con tus datos SMTP y preferencias.

4. Inicia la interfaz:
   ```bash
   python abrir_interfaz.py
   ```

## 📁 Estructura principal

- abrir_interfaz.py: arranque de la aplicación.
- interfaz_u4u.py: interfaz gráfica.
- enviar_correos.py: motor de envío y manejo de contactos.
- plantilla_u4u.html: plantilla activa.
- adjuntos/: archivos que se pueden adjuntar a los correos.
- logs/: salidas y vistas previas.

## 🔐 Importante

No subas datos reales de contactos, credenciales SMTP ni archivos sensibles. El proyecto ya incluye un .gitignore preparado para eso.

## 💡 Sugerencia

Antes de enviar lotes reales, ejecuta una prueba y apruébala desde la interfaz para activar el envío del batch.
