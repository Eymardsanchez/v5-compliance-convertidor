"""Análisis de resultados de IVA por periodo. Referencia: blueprint sección 12.

LIMITACIÓN DE ALCANCE (obligatoria): los acuses solo muestran el resultado neto
del periodo; no traen el renglón 'acreditamiento de saldo a favor de periodos
anteriores'. Por tanto el remanente acumulado de IVA NO es determinable solo con
acuses. La bandera `remanente_iva_determinable` es siempre False.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from .config import clasifica_concepto
from .models import Declaracion, ResultadoIVA

TEXTO_ALCANCE = (
    "ALCANCE IVA: los acuses solo muestran el resultado neto del periodo (a favor "
    "o a cargo); no incluyen el renglón interno 'acreditamiento de saldo a favor de "
    "periodos anteriores'. La serie por periodo NO acumula, por lo que el remanente "
    "de saldo a favor de IVA NO es determinable solo con acuses. Para determinarlo se "
    "requieren las declaraciones detalladas (vista previa), los papeles de trabajo de "
    "IVA (trasladado/acreditable/retenido por mes) o los acuses de devolución (FED)."
)


@dataclass
class AnalisisIVA:
    resultados: list[ResultadoIVA] = field(default_factory=list)
    meses_con_saldo_favor: int = 0
    suma_saldos_favor: float = 0.0
    suma_iva_a_cargo: float = 0.0
    remanente_iva_determinable: bool = False
    texto_alcance: str = TEXTO_ALCANCE
    compensaciones_detectadas: list = field(default_factory=list)


def analizar_iva(vigentes: list[Declaracion]) -> AnalisisIVA:
    an = AnalisisIVA()
    for d in vigentes:
        if d.ejercicio is None or d.periodo is None:
            continue
        for c in d.conceptos:
            if clasifica_concepto(c.nombre) != 'IVA':  # excluye IVA_RET
                continue
            a_favor = c.a_favor or 0.0
            a_cargo = c.a_cargo or 0.0
            an.resultados.append(ResultadoIVA(
                ejercicio=d.ejercicio,
                periodo=d.periodo,
                a_favor=a_favor,
                a_cargo=a_cargo,
                cantidad_pagar=c.cantidad_pagar or 0.0,
                tipo=d.tipo,
                operacion=d.operacion or "",
            ))
            if c.origen is not None:
                an.compensaciones_detectadas.append({
                    'ejercicio': d.ejercicio, 'periodo': d.periodo,
                    'tipo_origen': c.origen.tipo_origen,
                    'saldo_aplicar': c.origen.saldo_aplicar,
                })

    an.resultados.sort(key=lambda r: (r.ejercicio, r.periodo))
    an.meses_con_saldo_favor = sum(1 for r in an.resultados if r.a_favor > 0)
    an.suma_saldos_favor = round(sum(r.a_favor for r in an.resultados), 2)
    an.suma_iva_a_cargo = round(sum(r.a_cargo for r in an.resultados), 2)
    return an
