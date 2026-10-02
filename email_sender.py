# -*- coding: utf-8 -*-
"""
email_sender.py — Módulo de envío de correo multi-proveedor
============================================================
Soporta:
  • Gmail        smtp.gmail.com:587  STARTTLS  (contraseña de aplicación)
  • Outlook/Hotmail smtp-mail.outlook.com:587 STARTTLS (pwd normal o app pwd)
  • Yahoo        smtp.mail.yahoo.com:465  SSL   (contraseña de aplicación)
  • SMTP Custom  host/puerto/SSL configurables

Uso básico:
    from email_sender import enviar_correo, detectar_proveedor, PROVIDERS
    enviar_correo(cfg, "Asunto", "Cuerpo", ["dest@x.com"], adjuntos=["/ruta/a.pdf"])
"""

import smtplib
import ssl
import mimetypes
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders

# ── Catálogo de proveedores ───────────────────────────────────────────────────
PROVIDERS: dict[str, dict] = {
    "gmail": {
        "nombre":     "Gmail",
        "host_smtp":  "smtp.gmail.com",
        "puerto":     587,
        "ssl":        False,          # False = STARTTLS, True = SSL directo
        "dominios":   ["gmail.com", "googlemail.com"],
        "nota": (
            "Requiere contraseña de aplicación (no tu contraseña normal).\n"
            "Actívala en: myaccount.google.com → Seguridad → "
            "Verificación en 2 pasos → Contraseñas de aplicaciones."
        ),
    },
    "outlook": {
        "nombre":    "Outlook / Hotmail",
        "host_smtp": "smtp-mail.outlook.com",
        "puerto":    587,
        "ssl":       False,
        "dominios":  [
            "outlook.com", "outlook.es", "outlook.com.mx",
            "hotmail.com", "hotmail.es", "hotmail.com.mx",
            "live.com",    "live.com.mx", "msn.com",
        ],
        "nota": (
            "Usa tu contraseña habitual de Microsoft.\n"
            "Si tienes autenticación de 2 factores activa, genera una\n"
            "contraseña de aplicación en: account.microsoft.com → Seguridad."
        ),
    },
    "yahoo": {
        "nombre":    "Yahoo Mail",
        "host_smtp": "smtp.mail.yahoo.com",
        "puerto":    465,
        "ssl":       True,            # Yahoo requiere SSL en 465
        "dominios":  ["yahoo.com", "yahoo.com.mx", "yahoo.es", "ymail.com"],
        "nota": (
            "Requiere contraseña de aplicación.\n"
            "Actívala en: Yahoo → Mi cuenta → Seguridad → "
            "Contraseñas de aplicaciones."
        ),
    },
    "smtp": {
        "nombre":    "SMTP Personalizado",
        "host_smtp": "",
        "puerto":    587,
        "ssl":       False,
        "dominios":  [],
        "nota": (
            "Ingresa manualmente el host y puerto de tu servidor SMTP.\n"
            "Consulta a tu proveedor de correo si debes usar SSL (465) "
            "o STARTTLS (587)."
        ),
    },
}

# ── Detección automática de proveedor ────────────────────────────────────────
def detectar_proveedor(email: str) -> str:
    """
    Detecta el proveedor SMTP a partir del dominio del email.
    Retorna una clave de PROVIDERS ('gmail', 'outlook', 'yahoo', 'smtp').
    """
    dominio = email.split("@")[-1].lower().strip() if "@" in email else ""
    for nombre, cfg in PROVIDERS.items():
        if dominio in cfg["dominios"]:
            return nombre
    return "smtp"


def config_proveedor(proveedor: str) -> dict:
    """Devuelve host_smtp, puerto y ssl del proveedor dado."""
    p = PROVIDERS.get(proveedor, PROVIDERS["smtp"])
    return {"host_smtp": p["host_smtp"], "puerto": p["puerto"], "ssl": p["ssl"]}


# ── Envío de correo ───────────────────────────────────────────────────────────
def enviar_correo(
    cfg_email:    dict,
    asunto:       str,
    cuerpo:       str,
    destinatarios: list[str],
    cc:           list[str] | None = None,
    adjuntos:     list[str] | None = None,
    cuerpo_html:  str | None = None,
) -> None:
    """
    Envía un correo mediante el proveedor configurado.

    cfg_email debe contener:
        remitente   str   — dirección de quien envía
        password    str   — contraseña normal o de aplicación
        host_smtp   str   — host del servidor SMTP
        puerto      int   — puerto (587 STARTTLS / 465 SSL)
        ssl         bool  — True → SMTP_SSL, False → SMTP + STARTTLS

    Parámetros opcionales:
        cc          list  — destinatarios en copia
        adjuntos    list  — rutas absolutas de archivos a adjuntar
        cuerpo_html str   — versión HTML del cuerpo (si se omite, solo texto plano)
    """
    remitente = cfg_email["remitente"].strip()
    password  = cfg_email["password"].strip()
    host      = cfg_email.get("host_smtp", "").strip()
    puerto    = int(cfg_email.get("puerto", 587))
    usar_ssl  = bool(cfg_email.get("ssl", False))

    if not remitente or not password or not host:
        raise ValueError(
            "Configuración de correo incompleta.\n"
            "Verifica remitente, contraseña y host SMTP en ⚙ Configuración."
        )

    # Construir mensaje
    if cuerpo_html:
        msg = MIMEMultipart("alternative")
        msg.attach(MIMEText(cuerpo,      "plain", "utf-8"))
        msg.attach(MIMEText(cuerpo_html, "html",  "utf-8"))
    else:
        msg = MIMEMultipart()
        msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    msg["From"]    = remitente
    msg["To"]      = ", ".join(destinatarios)
    msg["Subject"] = asunto
    if cc:
        msg["Cc"] = ", ".join(cc)

    # Adjuntos
    for ruta in (adjuntos or []):
        if not os.path.isfile(ruta):
            continue
        tipo, _ = mimetypes.guess_type(ruta)
        if tipo is None:
            tipo = "application/octet-stream"
        tp, st = tipo.split("/", 1)
        with open(ruta, "rb") as f:
            parte = MIMEBase(tp, st)
            parte.set_payload(f.read())
        encoders.encode_base64(parte)
        parte.add_header(
            "Content-Disposition", "attachment",
            filename=os.path.basename(ruta),
        )
        msg.attach(parte)

    todos = list(destinatarios) + list(cc or [])

    # Conexión y envío
    tls = ssl.create_default_context()
    if usar_ssl:
        with smtplib.SMTP_SSL(host, puerto, timeout=30, context=tls) as server:
            server.login(remitente, password)
            server.sendmail(remitente, todos, msg.as_bytes())
    else:
        with smtplib.SMTP(host, puerto, timeout=30) as server:
            server.ehlo()
            server.starttls(context=tls)
            server.ehlo()
            server.login(remitente, password)
            server.sendmail(remitente, todos, msg.as_bytes())


def probar_conexion(cfg_email: dict) -> tuple[bool, str]:
    """
    Prueba la autenticación SMTP sin enviar ningún correo.
    Retorna (exito: bool, mensaje: str).
    """
    remitente = cfg_email.get("remitente", "").strip()
    password  = cfg_email.get("password",  "").strip()
    host      = cfg_email.get("host_smtp", "").strip()
    puerto    = int(cfg_email.get("puerto", 587))
    usar_ssl  = bool(cfg_email.get("ssl", False))

    try:
        tls = ssl.create_default_context()
        if usar_ssl:
            with smtplib.SMTP_SSL(host, puerto, timeout=15, context=tls) as server:
                server.login(remitente, password)
        else:
            with smtplib.SMTP(host, puerto, timeout=15) as server:
                server.ehlo()
                server.starttls(context=tls)
                server.ehlo()
                server.login(remitente, password)
        return True, f"✅ Conexión exitosa con {host}:{puerto}"
    except smtplib.SMTPAuthenticationError:
        return False, (
            "❌ Error de autenticación.\n"
            "Verifica tu correo y contraseña.\n"
            "Para Gmail/Yahoo se requiere contraseña de aplicación."
        )
    except smtplib.SMTPConnectError as e:
        return False, f"❌ No se pudo conectar a {host}:{puerto}\n{e}"
    except Exception as e:
        return False, f"❌ Error: {e}"
