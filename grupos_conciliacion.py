"""
grupos_conciliacion.py — Grupos de RFC para conciliación multi-RFC.

Permite agrupar varios RFCs bajo un nombre. En la conciliación (Fase 5),
los totales de todos los CFDIs de los RFCs del grupo se suman y se cruzan
contra el estado de cuenta como un solo movimiento bancario.

Persistencia: %APPDATA%\\ConvertidorBancos\\grupos_{RFC_empresa}.json
"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path


# ── Persistencia ──────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _path(rfc_empresa: str) -> Path:
    rfc_safe = rfc_empresa.strip().upper().replace(" ", "_")
    return _app_dir() / f"grupos_{rfc_safe}.json"


# ── Dataclass ─────────────────────────────────────────────────────────────────

@dataclass
class GrupoConciliacion:
    nombre:           str   = ""
    rfcs_proveedores: list  = field(default_factory=list)   # retiros (empresa paga)
    rfcs_clientes:    list  = field(default_factory=list)   # depósitos (empresa cobra)
    color_hex:        str   = "4527A0"
    activo:           bool  = True
    tolerancia:       float = 1.00   # MXN — margen para el cruce de sumas

    @property
    def rfcs(self) -> list:
        """Todos los RFCs del grupo (proveedores + clientes), sin duplicados."""
        seen: set = set()
        out: list = []
        for r in self.rfcs_proveedores + self.rfcs_clientes:
            if r not in seen:
                seen.add(r)
                out.append(r)
        return out


# ── API pública ───────────────────────────────────────────────────────────────

def cargar_grupos(rfc_empresa: str) -> list[GrupoConciliacion]:
    """Retorna los grupos guardados, o [] si aún no existe el archivo."""
    try:
        raw = json.loads(_path(rfc_empresa).read_text(encoding="utf-8"))
        result: list[GrupoConciliacion] = []
        for g in raw:
            # Migrar formato anterior que tenía 'rfcs' plano
            if "rfcs" in g and "rfcs_proveedores" not in g:
                g["rfcs_proveedores"] = g.pop("rfcs")
                g["rfcs_clientes"]    = []
            else:
                g.pop("rfcs", None)
                g.setdefault("rfcs_proveedores", [])
                g.setdefault("rfcs_clientes",    [])
            result.append(GrupoConciliacion(**g))
        return result
    except Exception:
        return []


def guardar_grupos(rfc_empresa: str, grupos: list[GrupoConciliacion]) -> None:
    _path(rfc_empresa).write_text(
        json.dumps([asdict(g) for g in grupos], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def agregar_grupo(rfc_empresa: str, grupo: GrupoConciliacion) -> None:
    grupos = cargar_grupos(rfc_empresa)
    grupos = [g for g in grupos if g.nombre != grupo.nombre]   # reemplazar si ya existe
    grupos.append(grupo)
    guardar_grupos(rfc_empresa, grupos)


def eliminar_grupo(rfc_empresa: str, nombre: str) -> None:
    grupos = [g for g in cargar_grupos(rfc_empresa) if g.nombre != nombre]
    guardar_grupos(rfc_empresa, grupos)


def actualizar_grupo(rfc_empresa: str, nombre_viejo: str,
                     grupo_nuevo: GrupoConciliacion) -> None:
    grupos = cargar_grupos(rfc_empresa)
    grupos = [grupo_nuevo if g.nombre == nombre_viejo else g for g in grupos]
    guardar_grupos(rfc_empresa, grupos)
