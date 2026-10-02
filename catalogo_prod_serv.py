"""
catalogo_prod_serv.py — Mapeo de ClaveProdServ SAT a cuentas contables.

Permite asignar cuentas de ingreso/gasto a claves o prefijos del catálogo
SAT de Productos y Servicios para la generación de Pólizas de Diario.

El prefijo puede ser de 2 a 8 dígitos. La búsqueda usa el prefijo más
largo que coincida con la clave del CFDI (más específico = mayor prioridad).

Persistencia: %APPDATA%\\ConvertidorBancos\\prod_serv_{RFC}.json
"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path


# ── Persistencia ──────────────────────────────────────────────────────────────

def _app_dir() -> Path:
    p = Path(os.environ.get("APPDATA", Path.home())) / "ConvertidorBancos"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _path(rfc_empresa: str) -> Path:
    return _app_dir() / f"prod_serv_{rfc_empresa.strip().upper()}.json"


# ── Dataclass ─────────────────────────────────────────────────────────────────

@dataclass
class ReglaProdServ:
    prefijo:        str = ""   # Prefijo SAT: "84", "841115", "84111500", etc.
    descripcion:    str = ""   # Etiqueta legible para el usuario
    cuenta_ingreso: str = ""   # Cuenta Cr Ingresos para CFDIs emitidas (Dr=CxC)
    cuenta_gasto:   str = ""   # Cuenta Dr Gastos para CFDIs recibidas (Cr=CxP)
    tipo:           str = ""   # "E"=Emitido/Ingreso, "R"=Recibido/Gasto, ""=Ambos


# ── API pública ───────────────────────────────────────────────────────────────

def cargar_reglas(rfc_empresa: str) -> list[ReglaProdServ]:
    """Retorna las reglas guardadas, o [] si no existe el archivo."""
    try:
        raw = json.loads(_path(rfc_empresa).read_text(encoding="utf-8"))
        return [ReglaProdServ(**r) for r in raw]
    except Exception:
        return []


def guardar_reglas(rfc_empresa: str, reglas: list[ReglaProdServ]) -> None:
    _path(rfc_empresa).write_text(
        json.dumps([asdict(r) for r in reglas], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def poblar_desde_cfdis(
    rfc_empresa: str,
    recibidos: list,
    emitidos: list,
    reglas_existentes: list[ReglaProdServ] | None = None,
) -> tuple[list[ReglaProdServ], int]:
    """
    Escanea los CFDIs e importa ClaveProdServ únicos como nuevas reglas.

    Distingue entre emitidos (tipo "E") y recibidos (tipo "R").
    Si el mismo código aparece en ambos se crean dos filas separadas.
    Sólo agrega códigos que aún no existan como (prefijo, tipo) en las reglas existentes.

    Retorna (lista_actualizada, n_nuevas).
    """
    if reglas_existentes is None:
        reglas_existentes = cargar_reglas(rfc_empresa)

    existentes = {(r.prefijo, r.tipo) for r in reglas_existentes}

    def _recolectar(cfdis: list) -> dict[str, str]:
        """Clave SAT → primera descripción encontrada."""
        out: dict[str, str] = {}
        for c in cfdis:
            clave = (getattr(c, "clave_prod_serv", "") or "").strip()
            if clave and clave not in out:
                out[clave] = (getattr(c, "descripcion_concepto", "") or "").strip()
        return out

    claves_e = _recolectar(emitidos)   # tipo E
    claves_r = _recolectar(recibidos)  # tipo R

    nuevas: list[ReglaProdServ] = []
    for clave in sorted(set(claves_e) | set(claves_r)):
        en_e = clave in claves_e
        en_r = clave in claves_r
        desc = claves_e.get(clave) or claves_r.get(clave, "")

        if en_e and en_r:
            # Aparece en ambos → crear ambas variantes si faltan
            for t in ("E", "R"):
                if (clave, t) not in existentes:
                    nuevas.append(ReglaProdServ(prefijo=clave, descripcion=desc, tipo=t))
                    existentes.add((clave, t))
        elif en_e:
            if (clave, "E") not in existentes:
                nuevas.append(ReglaProdServ(prefijo=clave, descripcion=desc, tipo="E"))
                existentes.add((clave, "E"))
        else:
            if (clave, "R") not in existentes:
                nuevas.append(ReglaProdServ(prefijo=clave, descripcion=desc, tipo="R"))
                existentes.add((clave, "R"))

    if nuevas:
        reglas_actualizadas = reglas_existentes + nuevas
        guardar_reglas(rfc_empresa, reglas_actualizadas)
        return reglas_actualizadas, len(nuevas)

    return reglas_existentes, 0


def cuenta_para_clave(
    clave:   str,
    tipo:    str,               # "ingreso" o "gasto"
    reglas:  list[ReglaProdServ],
    default: str = "",
) -> str:
    """
    Retorna la cuenta contable para una ClaveProdServ SAT.

    Busca por el prefijo más largo que coincida (más específico gana).
    Si ninguna regla aplica, retorna `default`.
    """
    clave = (clave or "").strip()
    best_len   = -1
    best_cuenta = default
    for r in reglas:
        if r.prefijo and clave.startswith(r.prefijo):
            if len(r.prefijo) > best_len:
                best_len    = len(r.prefijo)
                best_cuenta = (r.cuenta_ingreso if tipo == "ingreso"
                               else r.cuenta_gasto)
    return best_cuenta or default
