"""Constantes, mapeos y clasificación de conceptos.

Referencia: blueprint secciones 7 y 9. La clasificación por palabras clave
unifica los distintos nombres que el mismo impuesto recibe según el año y el
régimen (General vs. RESICO PM).
"""
from __future__ import annotations

MESES: dict[str, int] = {
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
    'julio': 7, 'agosto': 8, 'septiembre': 9, 'octubre': 10, 'noviembre': 11,
    'diciembre': 12,
}

MES_ABR = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep',
           'Oct', 'Nov', 'Dic']


def clasifica_concepto(nombre: str) -> str:
    """Clasifica un concepto de pago por palabras clave.

    Devuelve una de: ISR_PM, ISR_RET_SUELDOS, ISR_RET_ASIMILADOS, IVA_RET,
    IVA, OTRO. El nombre se normaliza a mayúsculas antes de comparar.
    """
    u = nombre.upper()
    if 'ISR' in u and 'PERSONAS MORALES' in u:
        # cubre "ISR personas morales" y "ISR simplificado ... personas morales"
        return 'ISR_PM'
    if 'ISR' in u and 'RETENC' in u and 'SALARIO' in u and 'ASIMIL' not in u:
        return 'ISR_RET_SUELDOS'
    if 'ISR' in u and 'RETENC' in u and 'ASIMIL' in u:
        return 'ISR_RET_ASIMILADOS'
    if 'IVA' in u and 'RETENC' in u:
        return 'IVA_RET'
    if 'IVA' in u or 'VALOR AGREGADO' in u:
        return 'IVA'
    return 'OTRO'


# Prioridad de complementarias (mayor peso supera al menor).
SUBTIPO_PESO: dict[str, int] = {
    'normal': 0,
    'complementaria': 1,
    'modificacion': 1,
    'no-presentada': 1,
    'dejar-sin-efecto': 2,
}
