"""
Bitácora de conversiones — ConvertidorBancos
=============================================
Registra cada conversión exitosa con: RFC, empresa, banco, periodo,
fecha/hora, usuario, número de movimientos y archivo generado.

Almacenamiento: %APPDATA%\\ConvertidorBancos\\bitacora.bin
Cifrado: XOR-SHA256 con MAC address (igual que usuarios.bin / empresas.bin).

API pública:
  registrar(entrada: dict)          -> None
  listar(rfc=None, banco=None,
         periodo=None)              -> list[dict]
  limpiar_anteriores_a(fecha_iso)   -> int   # registros eliminados
  total()                           -> int
"""

import hashlib
import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Optional


# ── Persistencia ──────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _path() -> Path:
    return _app_dir() / "bitacora.bin"

def _cipher_key() -> bytes:
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()

def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def _load() -> list:
    try:
        if not _path().exists():
            return []
        raw = _xor(_path().read_bytes(), _cipher_key())
        data = json.loads(raw.decode("utf-8"))
        return data if isinstance(data, list) else []
    except Exception:
        return []

def _save(data: list) -> None:
    _path().write_bytes(_xor(json.dumps(data, ensure_ascii=False).encode("utf-8"), _cipher_key()))


# ── API pública ───────────────────────────────────────────────────────────────

def registrar(
    rfc: str,
    empresa: str,
    banco: str,
    periodo: str,
    fecha_corte: str,
    movimientos: int,
    usuario: str,
    archivo_origen: str = "",
    archivo_excel: str = "",
) -> None:
    """
    Guarda una entrada en la bitácora.

    Parámetros
    ----------
    rfc            : RFC del titular de la cuenta (ej. "GAG250312RE0")
    empresa        : Nombre de la empresa (del catálogo)
    banco          : Nombre del banco detectado (ej. "BBVA")
    periodo        : Carpeta de mes generada (ej. "2026-03")
    fecha_corte    : Fecha de corte extraída del PDF (ej. "31/03/2026")
    movimientos    : Número de movimientos exportados
    usuario        : Usuario del sistema que realizó la conversión
    archivo_origen : Nombre del archivo PDF procesado
    archivo_excel  : Ruta completa del Excel generado
    """
    data = _load()
    data.append({
        "id":             str(uuid.uuid4())[:8],
        "fecha_hora":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario":        usuario,
        "rfc":            rfc.upper().strip(),
        "empresa":        empresa,
        "banco":          banco,
        "periodo":        periodo,           # "YYYY-MM"
        "fecha_corte":    fecha_corte,       # "dd/mm/yyyy"
        "movimientos":    movimientos,
        "archivo_origen": os.path.basename(archivo_origen),
        "archivo_excel":  archivo_excel,
    })
    _save(data)


def listar(
    rfc: Optional[str] = None,
    banco: Optional[str] = None,
    periodo: Optional[str] = None,
) -> List[dict]:
    """
    Retorna entradas ordenadas del más reciente al más antiguo.
    Filtra opcionalmente por RFC, banco y/o periodo (YYYY-MM).
    """
    data = _load()
    if rfc:
        rfc_upper = rfc.strip().upper()
        data = [e for e in data if e.get("rfc", "") == rfc_upper]
    if banco:
        data = [e for e in data if banco.lower() in e.get("banco", "").lower()]
    if periodo:
        data = [e for e in data if e.get("periodo", "") == periodo]
    return list(reversed(data))


def total() -> int:
    """Retorna el número total de conversiones registradas."""
    return len(_load())


def limpiar_anteriores_a(fecha_iso: str) -> int:
    """
    Elimina registros anteriores a fecha_iso (formato 'YYYY-MM-DD').
    Retorna el número de registros eliminados.
    """
    data    = _load()
    filtros = [e for e in data if e.get("fecha_hora", "") >= fecha_iso]
    eliminados = len(data) - len(filtros)
    if eliminados:
        _save(filtros)
    return eliminados
