"""
diagnostico_conciliacion.py — ¿POR QUÉ no concilió este movimiento?
====================================================================
Herramienta de diagnóstico. NO modifica la lógica de conciliacion.py; solo
la observa y explica, movimiento por movimiento, la razón exacta por la que
un cargo/depósito bancario quedó sin CFDI. Convierte "no concilia" en
"no concilia PORQUE ...".

Uso rápido:
    from conciliacion import conciliar_multibanco
    from diagnostico_conciliacion import diagnosticar_pendientes
    resultados = conciliar_multibanco(bancos, recibidos, emitidos)
    for banco in resultados:
        for linea in diagnosticar_pendientes(banco.sin_cfdi, recibidos, emitidos):
            print(linea)
"""

from __future__ import annotations

from conciliacion import (
    MovBanco, FORMAS_PAGO_BANCARIZADAS, _mes_fecha, _uuids_con_complemento,
)
from cfdi_parser import CFDI


def _c(x: float) -> int:
    return int(round(x * 100))


def diagnosticar_movimiento(
    mov:       MovBanco,
    recibidos: list[CFDI],
    emitidos:  list[CFDI],
    usados_r:  set[str] | None = None,
    usados_e:  set[str] | None = None,
) -> str:
    """Explica en español por qué este movimiento no encontró CFDI."""
    usados_r = usados_r or set()
    usados_e = usados_e or set()

    if mov.monto == 0:
        return "Sin monto (retiro y depósito en 0)."

    es_retiro = mov.retiro > 0
    pool   = recibidos if es_retiro else emitidos
    usados = usados_r if es_retiro else usados_e
    lado   = ("RECIBIDOS (retiro → pagó a proveedor/empleado)" if es_retiro
              else "EMITIDOS (depósito → le pagó un cliente)")

    if not pool:
        return (f"No hay CFDIs del lado {lado}. Revisa: ¿RFC de la empresa "
                f"correcto?, ¿cargaste la carpeta de XMLs de ese lado?")

    monto_c = _c(mov.monto)
    mes     = mov.mes or _mes_fecha(mov.fecha)
    con_comp = _uuids_con_complemento(recibidos + emitidos)

    # ── CFDIs con el MISMO monto exacto (cualquier mes) ──────────────────────
    mismo_monto = [c for c in pool if _c(c.total) == monto_c]

    if not mismo_monto:
        # ¿podría ser una SUMA de varias facturas del mismo mes?
        cand_mes = [c for c in pool if c.mes == mes and _c(c.total) < monto_c]
        suma_posible = len(cand_mes) >= 2 and sum(_c(c.total) for c in cand_mes) >= monto_c
        partes = [f"No hay ningún CFDI de {lado} con monto exacto ${mov.monto:,.2f}."]
        if suma_posible:
            partes.append("Podría ser SUMA de varias facturas del mes (cruce por suma).")
        partes.append("O el banco pagó NETO por retención (el total del CFDI difiere), "
                      "o simplemente no existe factura para ese movimiento.")
        return " ".join(partes)

    # ── Hay CFDIs con ese monto: diagnosticar cada uno ───────────────────────
    razones: list[str] = []
    for c in mismo_monto:
        et = f"CFDI {c.uuid[:8]} (mes {c.mes}, forma {c.forma_pago or '—'}, método {c.metodo_pago or '—'})"

        if c.uuid.lower() in usados:
            razones.append(f"{et}: YA se usó en otro movimiento (revisa duplicados de monto).")
            continue

        es_ppd = (c.metodo_pago == "PPD") or (c.forma_pago == "99")

        if c.mes != mes:
            if es_ppd:
                if c.uuid.lower() in con_comp:
                    razones.append(f"{et}: es de otro mes y es PPD CON complemento → debería cruzar "
                                   f"por complemento; revisa que el complemento tenga forma bancarizada.")
                else:
                    razones.append(f"{et}: es de otro mes, PPD 99 SIN complemento emitido → pendiente de complemento.")
            else:
                razones.append(f"{et}: es de mes {c.mes} y el banco es mes {mes}. "
                               f"PUE solo cruza en el MISMO mes (los meses distintos son solo para PPD).")
            continue

        # mismo mes
        if es_ppd:
            if c.uuid.lower() in con_comp:
                razones.append(f"{et}: PPD con complemento, mismo mes → debería cruzar; "
                               f"revisa que la forma del complemento sea bancarizada.")
            else:
                razones.append(f"{et}: PPD 99 SIN complemento emitido → pendiente de regularizar.")
            continue

        if c.forma_pago not in FORMAS_PAGO_BANCARIZADAS and c.tipo != "N":
            razones.append(f"{et}: forma de pago {c.forma_pago or '—'} NO es bancarizada "
                           f"(efectivo/compensación/etc.) → no se busca contra el banco.")
            continue

        razones.append(f"{et}: mismo monto, mismo mes y bancarizada → DEBERÍA cruzar por monto+mes. "
                       f"Si no cruzó, revisa que no esté ya usado ni duplicado con otro movimiento igual.")

    return "Hay CFDI(s) con ese monto:\n     - " + "\n     - ".join(razones)


def diagnosticar_pendientes(
    sin_cfdi:  list[MovBanco],
    recibidos: list[CFDI],
    emitidos:  list[CFDI],
) -> list[str]:
    """Una línea de diagnóstico por cada movimiento sin conciliar."""
    out: list[str] = []
    for mov in sin_cfdi:
        signo = "RETIRO " if mov.retiro > 0 else "DEPÓSITO"
        out.append(
            f"[{signo} ${mov.monto:,.2f} · {mov.fecha}] {mov.descripcion[:50]}\n"
            f"   → {diagnosticar_movimiento(mov, recibidos, emitidos)}"
        )
    return out
