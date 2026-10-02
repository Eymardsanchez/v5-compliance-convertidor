"""Deduplicación y determinación de la declaración vigente.

Referencia: blueprint sección 10.

Regla 1 — dedup por operación: el mismo acuse aparece con y sin año en el
nombre. Clave única = operacion; conservar la primera ocurrencia.

Regla 2 — vigente por (ejercicio, periodo): la de mayor fecha/hora de
presentación. Si la vigente es 'dejar-sin-efecto', el periodo queda sin efecto.

CRÍTICO: todo el análisis posterior usa SOLO las vigentes. Una complementaria
puede revertir una compensación previa.
"""
from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path

from .models import Declaracion


def _clave_dedup(d: Declaracion) -> str:
    if d.operacion:
        return f"op:{d.operacion}"
    # Sin número de operación: hash del nombre de archivo (y advertir aguas arriba).
    return "hash:" + hashlib.md5(Path(d.archivo).name.encode('utf-8')).hexdigest()


def deduplicar(declaraciones: list[Declaracion]) -> list[Declaracion]:
    """Conserva la primera ocurrencia por número de operación."""
    vistas: dict[str, Declaracion] = {}
    for d in declaraciones:
        clave = _clave_dedup(d)
        if clave not in vistas:
            vistas[clave] = d
    return list(vistas.values())


def _dt(d: Declaracion) -> datetime:
    return d.fecha_presentacion or datetime(1900, 1, 1)


def vigentes_mensuales(declaraciones: list[Declaracion]) -> dict[tuple[int, int], Declaracion]:
    """Devuelve {(ejercicio, periodo): declaracion vigente} para las mensuales."""
    grupos: dict[tuple[int, int], list[Declaracion]] = {}
    for d in declaraciones:
        if d.es_anual:
            continue
        if d.ejercicio is None or d.periodo is None:
            continue
        grupos.setdefault((d.ejercicio, d.periodo), []).append(d)
    return {k: max(g, key=_dt) for k, g in grupos.items()}


def anuales(declaraciones: list[Declaracion]) -> list[Declaracion]:
    """Devuelve las anuales vigentes (una por ejercicio, la de mayor fecha)."""
    grupos: dict[int, list[Declaracion]] = {}
    for d in declaraciones:
        if d.es_anual and d.ejercicio is not None:
            grupos.setdefault(d.ejercicio, []).append(d)
    return [max(g, key=_dt) for g in grupos.values()]


def esta_sin_efecto(d: Declaracion) -> bool:
    return d.subtipo == 'dejar-sin-efecto'
