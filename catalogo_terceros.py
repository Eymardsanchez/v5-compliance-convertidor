"""
catalogo_terceros.py — Catálogo de Clientes y Proveedores por empresa

Storage: %APPDATA%\\ConvertidorBancos\\terceros_{RFC}.json
Auto-población desde XMLs:
  • emitidos tipo I  → clientes  (Emisor = empresa, Receptor = cliente)
  • recibidos tipo I → proveedores (Receptor = empresa, Emisor = proveedor)

Asociación con cuenta contable: por similitud de nombre contra catálogo de cuentas.
"""

from __future__ import annotations

import json
import os
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from cfdi_parser import CFDI
    from catalogo_cuentas import CuentaContable


# ── Persistencia ──────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p

def _path(rfc_empresa: str) -> Path:
    rfc_safe = rfc_empresa.strip().upper().replace(" ", "_")
    return _app_dir() / f"terceros_{rfc_safe}.json"


# ── Dataclass ─────────────────────────────────────────────────────────────────

@dataclass
class Tercero:
    rfc:              str   # "AAA010101AAA"
    nombre:           str   # "Industrial Azteca SA de CV"
    tipo:             str   # "cliente" / "proveedor"
    cuenta_contable:  str   # "2110-001-000" (cuenta en catálogo COI)

    @property
    def label(self) -> str:
        icono = "👤" if self.tipo == "cliente" else "🏭"
        cta   = f" → {self.cuenta_contable}" if self.cuenta_contable else " → (sin cuenta)"
        return f"{icono} {self.rfc}  {self.nombre}{cta}"


# ── Helpers ───────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", str(s).upper())
        if not unicodedata.combining(c)
    ).strip()


def _similitud(a: str, b: str) -> float:
    """Fracción de palabras de 'a' que aparecen en 'b' (0-1)."""
    palabras_a = set(_norm(a).split())
    palabras_b = set(_norm(b).split())
    if not palabras_a:
        return 0.0
    comunes = palabras_a & palabras_b
    return len(comunes) / len(palabras_a)


# ── API pública ───────────────────────────────────────────────────────────────

def cargar_terceros(rfc_empresa: str) -> list[Tercero]:
    """Carga el catálogo de terceros. Retorna [] si no existe."""
    try:
        raw = json.loads(_path(rfc_empresa).read_text(encoding="utf-8"))
        return [Tercero(**r) for r in raw]
    except Exception:
        return []


def guardar_terceros(rfc_empresa: str, terceros: list[Tercero]) -> None:
    _path(rfc_empresa).write_text(
        json.dumps([asdict(t) for t in terceros], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def poblar_desde_cfdis(
    rfc_empresa: str,
    todos_recibidos: list["CFDI"],
    todos_emitidos:  list["CFDI"],
) -> tuple[int, int]:
    """
    Extrae clientes (emitidos tipo I) y proveedores (recibidos tipo I) de los XMLs.
    Agrega solo los terceros que no existen ya (por RFC).
    Retorna (nuevos_clientes, nuevos_proveedores).
    """
    existentes = cargar_terceros(rfc_empresa)
    rfcs_vistos = {t.rfc.upper() for t in existentes}

    nuevos_cli  = 0
    nuevos_prov = 0

    for cfdi in todos_emitidos:
        if cfdi.tipo != "I":
            continue
        rfc = cfdi.rfc_receptor.strip().upper()
        if not rfc or rfc in rfcs_vistos:
            continue
        existentes.append(Tercero(
            rfc=rfc,
            nombre=cfdi.nombre_receptor.strip() or rfc,
            tipo="cliente",
            cuenta_contable="",
        ))
        rfcs_vistos.add(rfc)
        nuevos_cli += 1

    for cfdi in todos_recibidos:
        if cfdi.tipo != "I":
            continue
        rfc = cfdi.rfc_emisor.strip().upper()
        if not rfc or rfc in rfcs_vistos:
            continue
        existentes.append(Tercero(
            rfc=rfc,
            nombre=cfdi.nombre_emisor.strip() or rfc,
            tipo="proveedor",
            cuenta_contable="",
        ))
        rfcs_vistos.add(rfc)
        nuevos_prov += 1

    guardar_terceros(rfc_empresa, existentes)
    return nuevos_cli, nuevos_prov


def auto_asociar_cuentas(
    rfc_empresa: str,
    cuentas: list["CuentaContable"],
    umbral: float = 0.6,
) -> int:
    """
    Para cada tercero sin cuenta_contable, busca la cuenta del catálogo cuyo
    nombre tiene mayor similitud con el nombre del tercero (>= umbral).
    Retorna la cantidad de asociaciones realizadas.
    """
    terceros = cargar_terceros(rfc_empresa)
    asociados = 0

    for t in terceros:
        if t.cuenta_contable:
            continue
        mejor_cta  = ""
        mejor_sim  = 0.0
        for c in cuentas:
            sim = _similitud(t.nombre, c.nombre)
            if sim > mejor_sim:
                mejor_sim = sim
                mejor_cta = c.cuenta
        if mejor_sim >= umbral:
            t.cuenta_contable = mejor_cta
            asociados += 1

    guardar_terceros(rfc_empresa, terceros)
    return asociados


def buscar_tercero(
    terceros: list[Tercero],
    rfc:      str = "",
    nombre:   str = "",
) -> Optional[Tercero]:
    """Busca por RFC exacto; si no hay match, por similitud de nombre."""
    rfc_u = rfc.strip().upper()
    if rfc_u:
        t = next((t for t in terceros if t.rfc == rfc_u), None)
        if t:
            return t
    if nombre:
        mejor, sim_max = None, 0.0
        for t in terceros:
            s = _similitud(nombre, t.nombre)
            if s > sim_max:
                sim_max = s
                mejor = t
        if sim_max >= 0.5:
            return mejor
    return None


def actualizar_cuenta(
    rfc_empresa: str,
    rfc_tercero: str,
    nueva_cuenta: str,
) -> bool:
    """Actualiza la cuenta_contable de un tercero. Retorna True si lo encontró."""
    terceros = cargar_terceros(rfc_empresa)
    rfc_u = rfc_tercero.strip().upper()
    for t in terceros:
        if t.rfc == rfc_u:
            t.cuenta_contable = nueva_cuenta.strip()
            guardar_terceros(rfc_empresa, terceros)
            return True
    return False
