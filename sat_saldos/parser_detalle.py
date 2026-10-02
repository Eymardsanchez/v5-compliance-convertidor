"""Parser de la DECLARACIÓN DETALLADA de IVA (vista previa del SAT).

A diferencia del acuse, la declaración detallada trae el renglón
'Acreditamiento del saldo a favor de periodos anteriores'. Con él se reconstruye
el remanente ACUMULADO real de IVA (que el acuse no permite determinar).

Estructura del PDF (validada contra los acuses reales 2025 de una PM RESICO):
- Documento multipágina; las páginas se separan con '\\f' (form feed).
- Cada página trae en su encabezado el impuesto: 'IVA simplificado de confianza'
  o 'ISR simplificado ...'. Los documentos pueden ser solo-IVA o combinados ISR+IVA.
- La página de DETERMINACIÓN de IVA contiene 'CANTIDAD A CARGO' y, según el
  resultado, 'IMPUESTO A CARGO' (a cargo) o 'IMPUESTO A FAVOR' (a favor).
- Los valores están alineados a la derecha; se toma el último número de la línea
  de la etiqueta. La primera ocurrencia de cada etiqueta es la de la determinación.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from .models import DetalleIVA
from .parser import _sin_acentos, num

MESMAP = {
    'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4, 'MAYO': 5, 'JUNIO': 6,
    'JULIO': 7, 'AGOSTO': 8, 'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11,
    'DICIEMBRE': 12,
}
_NUM = r'\(?-?\$?[\d,]+\.?\d*\)?'
_RE_MES = re.compile(r'\b(' + '|'.join(MESMAP) + r')\b')


def es_declaracion_detallada(texto: str) -> bool:
    """True si el texto corresponde a una declaración detallada de IVA."""
    u = texto.upper()
    return 'IVA ACREDITABLE DEL PERIODO' in u or 'ACREDITABLE DEL PERIODO' in u


def _pagina_determinacion_iva(texto: str) -> Optional[str]:
    """Devuelve la página de IVA con la determinación (a cargo / a favor)."""
    for p in texto.split('\f'):
        u = p.upper()
        if 'IVA SIMPLIFICADO' not in u and 'VALOR AGREGADO' not in u:
            continue
        if 'CANTIDAD A CARGO' in u and ('IMPUESTO A CARGO' in u or 'IMPUESTO A FAVOR' in u):
            return p
    return None


def _valor(pagina: str, etiqueta: str, occ: int = 0) -> Optional[float]:
    """Último número de la occ-ésima línea que hace match con la etiqueta."""
    lineas = [l for l in pagina.split('\n') if re.search(etiqueta, l, re.I)]
    if len(lineas) <= occ:
        return None
    nums = re.findall(_NUM, lineas[occ])
    return num(nums[-1]) if nums else None


def _periodo(texto: str) -> Optional[int]:
    """Periodo mensual: metadatos si están, si no el primer mes en páginas IVA."""
    flat = re.sub(r'[ \t]+', ' ', texto)
    m = re.search(r'Per[ií]odo de la declaraci[oó]n:\s*([A-Za-zÁÉÍÓÚáéíóú]+)', flat)
    if m and m.group(1).strip():
        return MESMAP.get(_sin_acentos(m.group(1)).upper())
    for p in texto.split('\f'):
        if 'IVA SIMPLIFICADO' in p.upper():
            mm = _RE_MES.search(_sin_acentos(p).upper())
            if mm:
                return MESMAP.get(mm.group(1))
    return None


def _meta(texto: str, archivo: str) -> tuple:
    flat = re.sub(r'[ \t]+', ' ', texto)
    rfc = re.search(r'RFC[:\s]+([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})', flat)
    ej = re.search(r'Ejercicio[:\s]+(\d{4})', flat)
    op = re.search(r'N[uú]mero de operaci[oó]n:\s*(\d+)', flat)
    fe = re.search(r'Fecha y hora de presentaci[oó]n:\s*(\d{2}/\d{2}/\d{4})\s*(\d{2}:\d{2})?', flat)
    dt = None
    if fe:
        try:
            dt = datetime.strptime(fe.group(1) + ' ' + (fe.group(2) or '00:00'), '%d/%m/%Y %H:%M')
        except ValueError:
            dt = None
    base = f"{Path(archivo).name} {flat}".lower()
    if 'dejar-sin-efecto' in base or 'dejar sin efecto' in base:
        subtipo = 'dejar-sin-efecto'
    elif 'modificaci' in base:
        subtipo = 'modificacion'
    elif 'complementaria' in base:
        subtipo = 'complementaria'
    else:
        subtipo = 'normal'
    return (rfc.group(1) if rfc else "",
            int(ej.group(1)) if ej else None,
            op.group(1) if op else None, dt, subtipo)


def parse_detalle_iva(texto: str, archivo: str = "") -> Optional[DetalleIVA]:
    """Parsea la determinación de IVA de una declaración detallada.

    `texto` debe ser salida de `pdftotext -layout` (conserva el form feed y la
    alineación). Devuelve None si no es una declaración detallada de IVA legible.
    """
    if not es_declaracion_detallada(texto):
        return None
    pagina = _pagina_determinacion_iva(texto)
    if pagina is None:
        return None

    rfc, ejercicio, operacion, dt, subtipo = _meta(texto, archivo)
    return DetalleIVA(
        archivo=archivo,
        rfc=rfc,
        ejercicio=ejercicio,
        periodo=_periodo(texto),
        operacion=operacion,
        fecha_presentacion=dt,
        subtipo=subtipo,
        a_cargo_periodo=_valor(pagina, r'^\s*CANTIDAD A CARGO', 0),
        a_favor_periodo=_valor(pagina, r'IMPUESTO A FAVOR', 0),
        acreditamiento_anterior=_valor(pagina, r'ACREDITAMIENTO DEL SALDO A FAVOR', 0) or 0.0,
        impuesto_a_cargo_final=_valor(pagina, r'^\s*IMPUESTO A CARGO', 0),
        cantidad_a_pagar=_valor(pagina, r'CANTIDAD A PAGAR', 0),
    )
