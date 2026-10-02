"""Análisis de saldos a favor de ISR: orígenes, aplicaciones y remanente.

Referencia: blueprint sección 11.

Clave de negocio:
- Los orígenes se identifican por (monto_original, fecha_saldo). El campo
  'Ejercicio' del bloque de origen a veces viene inconsistente en el acuse.
- El remanente nominal = monto_original - Σ saldo_aplicar(vigentes).
- El SAT actualiza por INPC (Art. 17-A CFF); el remanente actualizado del
  último acuse es informativo, no la cifra final.
- Solo se cuentan las declaraciones vigentes (una complementaria puede revertir
  una compensación).
"""
from __future__ import annotations

from typing import Optional

from .config import MES_ABR, clasifica_concepto
from .models import Concepto, Declaracion, OrigenSaldo


def _clave_origen(monto_original: Optional[float], fecha_saldo: Optional[str]) -> tuple:
    return (round(monto_original, 2) if monto_original is not None else None, fecha_saldo)


def _etiqueta_pago_indebido(d: Declaracion) -> str:
    if d.periodo and 1 <= d.periodo <= 12:
        return f"Pago indebido {MES_ABR[d.periodo].lower()}-{d.ejercicio}"
    return f"Pago indebido {d.ejercicio}"


def analizar_isr(
    vigentes: list[Declaracion],
    anuales: list[Declaracion],
) -> list[OrigenSaldo]:
    """Reconstruye los orígenes de saldo a favor de ISR y sus aplicaciones.

    `vigentes`: mensuales vigentes. `anuales`: anuales vigentes.
    """
    origenes: dict[tuple, OrigenSaldo] = {}

    # -- Paso A.1: orígenes desde las anuales con impuesto a favor > 0 --------
    for a in anuales:
        favor = next((c.a_favor for c in a.conceptos
                      if c.a_favor and clasifica_concepto(c.nombre) == 'ISR_PM'), None)
        if favor is None:
            favor = next((c.a_favor for c in a.conceptos if c.a_favor), None)
        if favor and favor > 0:
            fecha = (a.fecha_presentacion.strftime("%d/%m/%Y")
                     if a.fecha_presentacion else None)
            clave = _clave_origen(favor, fecha)
            origenes[clave] = OrigenSaldo(
                etiqueta=f"Anual {a.ejercicio}",
                ejercicio=a.ejercicio,
                monto_original=favor,
                fecha_origen=fecha,
                op_origen=a.operacion,
                estatus="Pendiente / sin aplicar",
            )

    # -- Paso A.2 + B: aplicaciones (compensaciones) en las vigentes ----------
    for d in vigentes:
        for c in d.conceptos:
            o = c.origen
            if o is None or not o.saldo_aplicar:
                continue
            clave = _clave_origen(o.monto_original, o.fecha_saldo)

            os_ = origenes.get(clave)
            if os_ is None:
                # Origen no visto por las anuales: crearlo desde el bloque.
                if o.tipo_origen.lower().startswith('pago'):
                    etiqueta = _etiqueta_pago_indebido(d)
                else:
                    etiqueta = (f"Anual {o.origen_ejercicio}"
                                if o.origen_ejercicio else "Saldo a favor")
                os_ = OrigenSaldo(
                    etiqueta=etiqueta,
                    ejercicio=o.origen_ejercicio,
                    monto_original=o.monto_original or 0.0,
                    fecha_origen=o.fecha_saldo,
                    op_origen=o.op_origen,
                )
                origenes[clave] = os_

            os_.aplicaciones.append({
                'periodo': f"{MES_ABR[d.periodo]}-{d.ejercicio}"
                           if d.periodo and 1 <= d.periodo <= 12 else str(d.ejercicio),
                'ejercicio': d.ejercicio,
                'mes': d.periodo,
                'fecha': (d.fecha_presentacion.strftime("%d/%m/%Y")
                          if d.fecha_presentacion else None),
                'tipo': d.tipo,
                'subtipo': d.subtipo,
                'isr_a_cargo': c.a_cargo,
                'compensado': o.saldo_aplicar,
                'remanente_act': o.remanente_act,
                'cantidad_pagar': c.cantidad_pagar,
                'operacion': d.operacion,
            })

    # -- Paso C: totales, remanente y estatus por origen ----------------------
    resultado: list[OrigenSaldo] = []
    for os_ in origenes.values():
        # Ordenar aplicaciones cronológicamente por (ejercicio, mes).
        os_.aplicaciones.sort(key=lambda x: (x.get('ejercicio') or 0, x.get('mes') or 0))
        total = sum(a['compensado'] for a in os_.aplicaciones if a['compensado'])
        os_.total_aplicado = round(total, 2)
        os_.remanente_nominal = round(os_.monto_original - total, 2)

        if os_.aplicaciones:
            ultimo = os_.aplicaciones[-1]
            ra = ultimo.get('remanente_act')
            comp = ultimo.get('compensado') or 0.0
            os_.remanente_actualizado_ultimo = (round(ra - comp, 2)
                                                if ra is not None else None)
            os_.estatus = ("Agotado" if abs(os_.remanente_nominal) < 0.01
                           else "Con remanente")
        else:
            os_.remanente_actualizado_ultimo = None
            os_.estatus = "Pendiente / sin aplicar"

        resultado.append(os_)

    # Orden estable de salida: anuales por ejercicio, luego pagos indebidos.
    resultado.sort(key=lambda o: (0 if o.etiqueta.startswith('Anual') else 1,
                                  o.ejercicio or 0, o.etiqueta))
    return resultado


def total_remanente_nominal(origenes: list[OrigenSaldo]) -> float:
    return round(sum(o.remanente_nominal for o in origenes), 2)
