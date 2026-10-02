"""
Registro de empresas — ConvertidorBancos
=========================================
Almacena RFC → { nombre, regimenes, regimenes_nombres } por empresa.
El campo regimenes contiene la lista de claves SAT (ej. ["626"]).

Formato RFC valido (SAT Mexico):
  Persona Moral : AAA000000AA0  (12 chars)
  Persona Fisica: AAAA000000AA0 (13 chars)

Almacenamiento (nuevo): **FUENTE ÚNICA en la FDB (Firebird), tabla EMPRESA** —
compartida por TODOS los módulos y pestañas. Respaldo/local: %APPDATA%\\
ConvertidorBancos\\empresas.bin (XOR-SHA256 con MAC), que además sirve de
fallback si Firebird no está disponible. La primera vez que la FDB responde y su
tabla EMPRESA está vacía, se migra automáticamente el contenido del .bin.

API publica (sin cambios — todos los módulos ya la usan):
  agregar_empresa(rfc, nombre)               -> (bool, str)
  actualizar_empresa(rfc, nombre)            -> (bool, str)
  registrar_desde_cif(pdf_path)             -> (bool, str, DatosCSF|None)
  actualizar_regimenes(rfc, codigos, nombres)-> (bool, str)
  obtener_regimenes(rfc)                     -> list[str]   # claves SAT
  obtener_empresa(rfc)                       -> dict | None
  eliminar_empresa(rfc)                      -> (bool, str)
  buscar_por_rfc(rfc)                        -> str | None
  listar_empresas()                          -> list[tuple]
  validar_rfc(rfc)                           -> bool
"""

import hashlib
import json
import os
import re
import uuid
from pathlib import Path
from typing import List, Optional, Tuple


# RFC SAT: 3-4 letras, 6 digitos, 3 alfanumericos (homoclave)
_RFC_RE = re.compile(
    r"^[A-Z&\xd1]{3,4}"   # 3-4 letras (incluyendo Ñ y &)
    r"\d{6}"                # 6 digitos de fecha AAMMDD
    r"[A-Z0-9]{3}$",        # 3 alfanumericos (homoclave)
    re.IGNORECASE,
)


# ── Respaldo local cifrado (.bin) — fallback + espejo ─────────────────────────
def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _path() -> Path:
    return _app_dir() / "empresas.bin"

def _cipher_key() -> bytes:
    return hashlib.sha256(str(uuid.getnode()).encode()).digest()

def _xor(data: bytes, key: bytes) -> bytes:
    return bytes(b ^ key[i % len(key)] for i, b in enumerate(data))

def _load_bin() -> dict:
    """dict {rfc: {"nombre", "regimenes", "regimenes_nombres"}} desde el .bin local.
    Migra el formato antiguo {rfc: nombre_str}."""
    try:
        if not _path().exists():
            return {}
        raw  = _xor(_path().read_bytes(), _cipher_key())
        data = json.loads(raw.decode("utf-8"))
        migrado = {}
        for rfc, val in data.items():
            if isinstance(val, str):
                migrado[rfc] = {"nombre": val, "regimenes": [], "regimenes_nombres": []}
            else:
                migrado[rfc] = val
        return migrado
    except Exception:
        return {}

def _save_bin(data: dict) -> None:
    try:
        _path().write_bytes(
            _xor(json.dumps(data, ensure_ascii=False).encode("utf-8"), _cipher_key()))
    except Exception:
        pass


# ── FUENTE ÚNICA: FDB (Firebird), tabla EMPRESA ───────────────────────────────
_persist = None          # instancia cacheada de FirebirdPersistencia
_migrado = False         # migración .bin → FDB ya intentada en este proceso


def _get_persist():
    """Devuelve una FirebirdPersistencia lista, o None si Firebird no responde.
    Cachea la conexión; si una operación falla, se invalida y se reintenta luego."""
    global _persist
    if _persist is not None:
        return _persist
    try:
        from persistencia import FirebirdPersistencia
        p = FirebirdPersistencia()
        p.inicializar()
        _persist = p
        _migrar_si_necesario(p)
        return p
    except Exception:
        return None


def _reset_persist():
    global _persist
    try:
        if _persist is not None:
            _persist.cerrar()
    except Exception:
        pass
    _persist = None


def _migrar_si_necesario(p) -> None:
    """Si la tabla EMPRESA está vacía y hay datos en el .bin local, los importa."""
    global _migrado
    if _migrado:
        return
    _migrado = True
    try:
        if p.leer_empresas():
            return                        # la FDB ya tiene datos → no migrar
        bind = _load_bin()
        for rfc, e in bind.items():
            p.guardar_empresa(rfc, e.get("nombre", ""),
                              e.get("regimenes", []), e.get("regimenes_nombres", []))
    except Exception:
        pass


def _load() -> dict:
    """dict {rfc: entry}. Fuente: FDB si responde; si no, respaldo local .bin."""
    p = _get_persist()
    if p is not None:
        try:
            filas = p.leer_empresas()
            return {e["rfc"]: {"nombre": e["nombre"], "regimenes": e["regimenes"],
                               "regimenes_nombres": e["regimenes_nombres"]}
                    for e in filas}
        except Exception:
            _reset_persist()
    return _load_bin()


def _guardar_entry(rfc: str, entry: dict) -> None:
    """Escribe una empresa en la FDB (si responde) y SIEMPRE en el respaldo .bin."""
    p = _get_persist()
    if p is not None:
        try:
            p.guardar_empresa(rfc, entry.get("nombre", ""),
                              entry.get("regimenes", []),
                              entry.get("regimenes_nombres", []))
        except Exception:
            _reset_persist()
    # Espejo local (fallback siempre fresco)
    bind = _load_bin()
    bind[rfc] = entry
    _save_bin(bind)


def _eliminar_entry(rfc: str) -> None:
    p = _get_persist()
    if p is not None:
        try:
            p.eliminar_empresa(rfc)
        except Exception:
            _reset_persist()
    bind = _load_bin()
    if rfc in bind:
        del bind[rfc]
        _save_bin(bind)


# ── API publica ───────────────────────────────────────────────────────────────

def validar_rfc(rfc: str) -> bool:
    """True si el RFC tiene el formato valido del SAT."""
    return bool(_RFC_RE.match(rfc.strip().upper()))


def agregar_empresa(rfc: str, nombre: str) -> Tuple[bool, str]:
    """Agrega o actualiza una empresa (sin modificar regímenes existentes)."""
    rfc    = rfc.strip().upper().replace(" ", "").replace("-", "")
    nombre = nombre.strip()

    if not rfc:
        return False, "El RFC es obligatorio."
    if not nombre:
        return False, "El nombre de la empresa es obligatorio."
    if not validar_rfc(rfc):
        return False, (
            f"RFC invalido: '{rfc}'.\n"
            "Formato: AAA000000AA0 (moral) o AAAA000000AA0 (fisica)."
        )

    data   = _load()
    accion = "actualizada" if rfc in data else "registrada"
    entry  = data.get(rfc, {"nombre": "", "regimenes": [], "regimenes_nombres": []})
    entry["nombre"] = nombre
    _guardar_entry(rfc, entry)
    return True, f"Empresa '{nombre}' {accion} con RFC {rfc}."


def registrar_desde_cif(pdf_path: str) -> Tuple[bool, str, Optional[object]]:
    """
    Parsea un PDF de Constancia de Situación Fiscal y guarda (o actualiza)
    la empresa con sus regímenes fiscales. Retorna (exito, mensaje, DatosCSF|None).
    """
    from cif_parser import parsear_cif
    datos = parsear_cif(pdf_path)
    if not datos:
        return False, "No se pudo leer el archivo como Constancia de Situación Fiscal.", None

    rfc    = datos.rfc
    nombre = datos.razon_social or rfc

    if not validar_rfc(rfc):
        return False, f"RFC extraido inválido: '{rfc}'.", None

    data  = _load()
    entry = data.get(rfc, {"nombre": "", "regimenes": [], "regimenes_nombres": []})
    entry["nombre"]             = nombre
    entry["regimenes"]          = datos.regimenes_codigos
    entry["regimenes_nombres"]  = datos.regimenes_nombres
    _guardar_entry(rfc, entry)

    regs_txt = ", ".join(
        f"{n} ({c})" if c else n
        for n, c in zip(datos.regimenes_nombres, datos.regimenes_codigos)
    ) or "(régimen no identificado)"
    return True, (
        f"Empresa registrada desde CSF:\n"
        f"  RFC: {rfc}\n"
        f"  Nombre: {nombre}\n"
        f"  Régimen(es): {regs_txt}"
    ), datos


def actualizar_regimenes(rfc: str, codigos: list, nombres: list) -> Tuple[bool, str]:
    """Actualiza manualmente los regímenes de una empresa ya registrada."""
    rfc  = rfc.strip().upper()
    data = _load()
    if rfc not in data:
        return False, f"RFC '{rfc}' no registrado. Usa agregar_empresa primero."
    entry = data[rfc]
    entry["regimenes"]         = codigos
    entry["regimenes_nombres"] = nombres
    _guardar_entry(rfc, entry)
    return True, f"Regímenes actualizados para {rfc}."


def obtener_regimenes(rfc: str) -> List[str]:
    """Retorna la lista de claves SAT de régimen para el RFC dado. [] si no está."""
    if not rfc:
        return []
    entry = _load().get(rfc.strip().upper())
    if not entry:
        return []
    return entry.get("regimenes", [])


def obtener_empresa(rfc: str) -> Optional[dict]:
    """Retorna el dict completo de la empresa o None (nombre, regimenes, regimenes_nombres)."""
    if not rfc:
        return None
    entry = _load().get(rfc.strip().upper())
    if not entry:
        return None
    if isinstance(entry, str):
        return {"nombre": entry, "regimenes": [], "regimenes_nombres": []}
    return entry


def actualizar_empresa(rfc: str, nombre: str) -> Tuple[bool, str]:
    """Actualiza el nombre de una empresa existente (conserva sus regímenes)."""
    rfc = rfc.strip().upper()
    data = _load()
    if rfc not in data:
        return False, f"RFC '{rfc}' no registrado. Usa agregar_empresa primero."
    entry = data[rfc]
    entry["nombre"] = nombre.strip()
    _guardar_entry(rfc, entry)
    return True, "Empresa actualizada correctamente."


def eliminar_empresa(rfc: str) -> Tuple[bool, str]:
    """Elimina una empresa del catalogo."""
    rfc  = rfc.strip().upper()
    data = _load()
    if rfc not in data:
        return False, f"RFC '{rfc}' no encontrado en el catalogo."
    nombre = data[rfc].get("nombre", rfc) if isinstance(data[rfc], dict) else str(data[rfc])
    _eliminar_entry(rfc)
    return True, f"Empresa '{nombre}' eliminada."


def buscar_por_rfc(rfc: str) -> Optional[str]:
    """Retorna el nombre de la empresa o None si no está registrada."""
    if not rfc:
        return None
    entry = _load().get(rfc.strip().upper())
    if not entry:
        return None
    return entry.get("nombre", "") if isinstance(entry, dict) else str(entry)


def listar_empresas() -> List[tuple]:
    """Lista [(rfc, nombre), ...] ordenada alfabeticamente por nombre."""
    data = _load()
    items = [(k, v.get("nombre", k) if isinstance(v, dict) else str(v))
             for k, v in data.items()]
    return sorted(items, key=lambda x: x[1].lower())
