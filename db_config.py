# -*- coding: utf-8 -*-
"""
db_config.py — Configuración de conexión a la base de datos (Firebird).

El cliente puede tener host/ruta/usuario/contraseña distintos, así que estos
parámetros NO deben ir fijos en el código: se guardan en un archivo JSON local y
se editan desde el GUI (Configuración → Base de datos).

Ubicación del archivo: ``%PROGRAMDATA%\\ConvertidorBancos\\db_config.json``
(la misma carpeta donde vive la BD por defecto; accesible al servicio Firebird,
fuera de OneDrive).

Parámetros (estándar de conexión Firebird):
  motor     : "firebird"
  host      : servidor (localhost si es local)
  port      : puerto del servicio (3050 por defecto)
  database  : ruta del archivo .fdb
  user      : usuario (SYSDBA por defecto)
  password  : contraseña
  charset   : juego de caracteres (UTF8)
  fbclient  : ruta a fbclient.dll (vacío = autodetección de 64-bit)
"""

from __future__ import annotations

import json
import os


def _carpeta_datos() -> str:
    base = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
    carpeta = os.path.join(base, "ConvertidorBancos")
    os.makedirs(carpeta, exist_ok=True)
    return carpeta


RUTA_CONFIG = os.path.join(_carpeta_datos(), "db_config.json")


def _fbclient_auto() -> str:
    for c in (r"C:\Windows\System32\fbclient.dll",
              r"C:\Program Files\Firebird\Firebird_5_0\fbclient.dll",
              r"C:\Program Files\Firebird\Firebird_5_0\bin\fbclient.dll",
              r"C:\Program Files\Firebird\Firebird_3_0\bin\fbclient.dll"):
        if os.path.exists(c):
            return c
    return ""


def defaults() -> dict:
    return {
        "motor":    "firebird",
        "host":     "localhost",
        "port":     3050,
        "database": os.path.join(_carpeta_datos(), "conciliacion.fdb"),
        "user":     "SYSDBA",
        "password": "",
        "charset":  "UTF8",
        "fbclient": "",           # vacío = autodetección
    }


def configurada() -> bool:
    """True si ya existe configuración de BD guardada; el GUI debe exigir configurarla antes de usar la app si es False."""
    return os.path.exists(RUTA_CONFIG)


def cargar() -> dict:
    """Devuelve la config (defaults combinados con lo guardado)."""
    cfg = defaults()
    try:
        if os.path.exists(RUTA_CONFIG):
            with open(RUTA_CONFIG, "r", encoding="utf-8") as fh:
                cfg.update({k: v for k, v in json.load(fh).items() if v is not None})
    except Exception:
        pass
    try:
        cfg["port"] = int(cfg.get("port") or 3050)
    except Exception:
        cfg["port"] = 3050
    return cfg


def guardar(cfg: dict) -> None:
    data = defaults()
    data.update(cfg or {})
    try:
        data["port"] = int(data.get("port") or 3050)
    except Exception:
        data["port"] = 3050
    with open(RUTA_CONFIG, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


def dsn(cfg: dict) -> str:
    """Arma el DSN de Firebird. Con puerto: ``host/port:ruta``; sin host (vacío)
    usa modo embebido con solo la ruta."""
    host = (cfg.get("host") or "").strip()
    db   = cfg.get("database") or ""
    if not host:
        return db                      # embebido
    port = int(cfg.get("port") or 3050)
    return f"{host}/{port}:{db}" if port and port != 0 else f"{host}:{db}"


def probar_conexion(cfg: dict) -> tuple[bool, str]:
    """Intenta conectar (creando la BD y el esquema si no existen). Devuelve
    (ok, mensaje). NO deja conexiones abiertas."""
    try:
        import persistencia as _p
        p = _p.FirebirdPersistencia(
            ruta_fdb=cfg.get("database"),
            fbclient=(cfg.get("fbclient") or None),
            host=cfg.get("host") or "localhost",
            user=cfg.get("user") or "SYSDBA",
            password=cfg.get("password") or "",
            port=int(cfg.get("port") or 3050),
            charset=cfg.get("charset") or "UTF8",
        )
        p.inicializar()
        p.cerrar()
        return True, ("Conexión correcta. La base y sus tablas están listas.\n"
                      f"Servidor: {cfg.get('host')}:{cfg.get('port')}\n"
                      f"BD: {cfg.get('database')}")
    except Exception as e:
        return False, f"No se pudo conectar:\n{type(e).__name__}: {e}"
