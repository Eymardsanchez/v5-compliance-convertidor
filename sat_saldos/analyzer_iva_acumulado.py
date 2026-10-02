"""Reconstrucción del remanente ACUMULADO de saldo a favor de IVA.

A partir de las declaraciones detalladas (vista previa), que sí traen el renglón
'Acreditamiento del saldo a favor de periodos anteriores':

    remanente[m] = remanente[m-1] + saldo_a_favor_generado[m] − acreditamiento[m]

Se parte de un saldo inicial (arrastre del ejercicio anterior). Si con saldo
inicial = 0 el remanente se vuelve negativo en algún mes, ese déficit revela que
EXISTE un saldo a favor inicial arrastrado del ejercicio previo de, al menos,
la magnitud del déficit máximo (no puede haber remanente negativo — Art. 6 LIVA).
"""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from .config import MES_ABR
from .models import DetalleIVA


@dataclass
class RenglonAcumulado:
    ejercicio: int
    periodo: int
    a_favor_generado: float
    acreditamiento: float
    a_cargo_periodo: float
    remanente_desde_cero: float   # remanente si el saldo inicial fuera 0 (puede ser negativo)
    operacion: Optional[str]
    subtipo: str
    fecha_presentacion: Optional[datetime] = None


@dataclass
class CapaOrigen:
    """Capa PEPS: saldo a favor generado en un mes y cuánto sobrevive al cierre."""
    origen: str          # "Jul-2025" o "Saldo inicial (previo)"
    generado: float
    restante: float
    operacion: Optional[str] = None
    fecha_presentacion: Optional[datetime] = None


@dataclass
class Consumo:
    """Un acreditamiento y de qué mes(es) de origen se tomó (PEPS)."""
    mes: str
    detalle: list         # list[(origen, monto)]
    faltante: float = 0.0 # lo que no alcanzó a cubrir el saldo disponible


@dataclass
class AnalisisIVAAcumulado:
    renglones: list[RenglonAcumulado] = field(default_factory=list)
    total_generado: float = 0.0
    total_acreditado: float = 0.0
    neto: float = 0.0                       # generado − acreditado (cambio del acumulado en el periodo)
    deficit_maximo: float = 0.0             # magnitud del punto más negativo desde cero
    saldo_inicial_requerido: float = 0.0    # ≥ este valor debió venir del ejercicio anterior
    cobertura: list[int] = field(default_factory=list)
    completo: bool = False                  # 12 meses presentes
    ejercicio: Optional[int] = None
    nota: str = ""
    # Trazabilidad PEPS del origen del saldo a favor.
    saldo_inicial: float = 0.0
    capas_cierre: list = field(default_factory=list)   # list[CapaOrigen] con restante > 0
    consumos: list = field(default_factory=list)       # list[Consumo]


def _dt(d: DetalleIVA) -> datetime:
    return d.fecha_presentacion or datetime(1900, 1, 1)


def vigentes_detalle(detalles: list[DetalleIVA]) -> dict[tuple[int, int], DetalleIVA]:
    """Dedup por operación y vigente por (ejercicio, periodo) = mayor fecha."""
    byop: dict[str, DetalleIVA] = {}
    for d in detalles:
        clave = d.operacion or d.archivo
        byop.setdefault(clave, d)
    grupos: dict[tuple[int, int], list[DetalleIVA]] = {}
    for d in byop.values():
        if d.ejercicio is None or d.periodo is None:
            continue
        grupos.setdefault((d.ejercicio, d.periodo), []).append(d)
    return {k: max(g, key=_dt) for k, g in grupos.items()}


def _trazar_peps(renglones: list[RenglonAcumulado], saldo_inicial: float
                 ) -> tuple[list[CapaOrigen], list[Consumo]]:
    """Traza el origen del saldo a favor por capas PEPS (lo más antiguo se aplica primero).

    Devuelve (capas_vivas_al_cierre, consumos). Cada capa indica de qué mes de
    origen proviene el saldo a favor que sobrevive; cada consumo indica de qué
    mes(es) salió cada acreditamiento.
    """
    capas: deque = deque()  # [CapaOrigen]
    if saldo_inicial and saldo_inicial > 0:
        capas.append(CapaOrigen("Saldo inicial (previo)", saldo_inicial, saldo_inicial))
    consumos: list[Consumo] = []
    for r in renglones:
        etq = etiqueta_mes(r.ejercicio, r.periodo)
        if r.a_favor_generado > 0:
            capas.append(CapaOrigen(etq, r.a_favor_generado, r.a_favor_generado,
                                    operacion=r.operacion,
                                    fecha_presentacion=r.fecha_presentacion))
        if r.acreditamiento > 0:
            pend = r.acreditamiento
            detalle = []
            while pend > 1e-6 and capas:
                cap = capas[0]
                toma = min(cap.restante, pend)
                cap.restante -= toma
                pend -= toma
                detalle.append((cap.origen, round(toma, 2)))
                if cap.restante <= 1e-6:
                    capas.popleft()
            consumos.append(Consumo(etq, detalle, round(pend, 2) if pend > 1e-6 else 0.0))
    vivas = [CapaOrigen(c.origen, round(c.generado, 2), round(c.restante, 2),
                        operacion=c.operacion, fecha_presentacion=c.fecha_presentacion)
             for c in capas]
    return vivas, consumos


def analizar_iva_acumulado(detalles: list[DetalleIVA],
                           saldo_inicial: float = 0.0) -> AnalisisIVAAcumulado:
    an = AnalisisIVAAcumulado()
    an.saldo_inicial = saldo_inicial or 0.0
    vig = vigentes_detalle(detalles)
    if not vig:
        an.nota = "Sin declaraciones detalladas de IVA legibles."
        return an

    an.ejercicio = max(ej for (ej, _) in vig)
    orden = sorted(vig.items(), key=lambda kv: kv[0])  # por (ejercicio, periodo)

    rem = 0.0
    peor = 0.0
    for (ej, per), d in orden:
        gen = d.a_favor_periodo or 0.0
        acr = d.acreditamiento_anterior or 0.0
        rem = rem + gen - acr
        peor = min(peor, rem)
        an.renglones.append(RenglonAcumulado(
            ejercicio=ej, periodo=per,
            a_favor_generado=gen, acreditamiento=acr,
            a_cargo_periodo=d.a_cargo_periodo or 0.0,
            remanente_desde_cero=round(rem, 2),
            operacion=d.operacion, subtipo=d.subtipo,
            fecha_presentacion=d.fecha_presentacion,
        ))
        an.total_generado += gen
        an.total_acreditado += acr

    an.total_generado = round(an.total_generado, 2)
    an.total_acreditado = round(an.total_acreditado, 2)
    an.neto = round(an.total_generado - an.total_acreditado, 2)
    an.deficit_maximo = round(-peor, 2) if peor < 0 else 0.0
    an.saldo_inicial_requerido = max(round(-peor, 2), 0.0)
    an.cobertura = sorted({per for (_, per) in vig})
    an.completo = len(an.cobertura) == 12

    # Trazabilidad PEPS: si no se capturó saldo inicial, se usa el mínimo requerido
    # para que las capas cierren (no puede haber consumo sin saldo disponible).
    base_peps = an.saldo_inicial if an.saldo_inicial > 0 else an.saldo_inicial_requerido
    an.capas_cierre, an.consumos = _trazar_peps(an.renglones, base_peps)

    if an.deficit_maximo > 0:
        an.nota = (
            f"El remanente calculado desde cero se vuelve negativo (hasta "
            f"−${an.deficit_maximo:,.2f}). Como no puede existir remanente negativo, "
            f"hubo un SALDO A FAVOR INICIAL arrastrado del ejercicio anterior de al menos "
            f"${an.saldo_inicial_requerido:,.2f}. El remanente al cierre = saldo inicial "
            f"(del año previo) + (${an.neto:,.2f}). Para el saldo inicial exacto se "
            f"requieren las declaraciones detalladas del ejercicio {(an.ejercicio or 1) - 1}."
        )
    else:
        an.nota = (
            f"Remanente acumulado del periodo (saldo inicial + neto ${an.neto:,.2f}). "
            f"Confirmar saldo inicial con el cierre del ejercicio anterior."
        )
    return an


def etiqueta_mes(ejercicio: int, periodo: int) -> str:
    return f"{MES_ABR[periodo]}-{ejercicio}" if 1 <= periodo <= 12 else f"{periodo}-{ejercicio}"
