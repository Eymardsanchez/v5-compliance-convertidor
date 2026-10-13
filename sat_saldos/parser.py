"""Parsing de texto de acuse -> Declaracion. Referencia: blueprint sección 9.

Los patrones aquí fueron validados contra 184 acuses reales. El texto se
'aplana' (colapsa espacios/tabs) antes de aplicar la mayoría de los regex,
porque `pdftotext -layout` deja alineación en columnas que estorba al parsear
etiqueta/valor.
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Optional

from .config import MESES
from .models import Concepto, Declaracion, OrigenCompensacion


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
def _flat(txt: str) -> str:
    """Colapsa espacios/tabs (no saltos de línea) para facilitar los regex."""
    return re.sub(r'[ \t]+', ' ', txt)


def num(s: Optional[str]) -> Optional[float]:
    """Convierte '1,234.56', '$1,234', '(1,388)' -> float. '(x)' es negativo."""
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    neg = s.startswith('(') and s.endswith(')')
    s = s.strip('()').replace('$', '').replace(',', '').strip()
    if s in ('', '-', '--'):
        return None
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _search(pat: str, texto: str, flags: int = 0) -> Optional[str]:
    m = re.search(pat, texto, flags)
    return m.group(1).strip() if m else None


# --------------------------------------------------------------------------- #
# Metadatos
# --------------------------------------------------------------------------- #
RE_RFC = r'RFC[:\s]+([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})'
RE_RAZON = r'(?:Denominaci[oó]n o raz[oó]n social)[:\s]+([^\n]+)'
RE_EJERCICIO = r'Ejercicio[:\s]+(\d{4})'
RE_PERIODO = r'Per[ií]odo de la declaraci[oó]n[:\s]+([A-Za-zÁÉÍÓÚáéíóú]+)'
RE_TIPO = r'Tipo de declaraci[oó]n[:\s]+([A-Za-zÁÉÍÓÚáéíóú]+)'
RE_FECHA = r'presentaci[oó]n[:\s]+(\d{2}/\d{2}/\d{4})\s*(\d{2}:\d{2})?'
RE_OPERACION = r'N[uú]mero de operaci[oó]n[:\s]+(\d+)'
RE_VERSION = r'Versi[oó]n[:\s]+(\d+)'


def _subtipo(nombre_archivo: str, flat: str) -> str:
    """Infiere el subtipo desde el nombre de archivo o del texto."""
    base = f"{nombre_archivo} {flat}".lower()
    if 'dejar-sin-efecto' in base or 'dejar sin efecto' in base:
        return 'dejar-sin-efecto'
    if 'obligacion-no-presentada' in base or 'obligación no presentada' in base \
            or 'obligacion no presentada' in base:
        return 'no-presentada'
    if 'modificacion' in base or 'modificación' in base:
        return 'modificacion'
    if 'complementaria' in base:
        return 'complementaria'
    return 'normal'


def _fecha_dt(flat: str) -> Optional[datetime]:
    m = re.search(RE_FECHA, flat)
    if not m:
        return None
    fecha, hora = m.group(1), m.group(2)
    try:
        if hora:
            return datetime.strptime(f"{fecha} {hora}", "%d/%m/%Y %H:%M")
        return datetime.strptime(fecha, "%d/%m/%Y")
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# Conceptos mensuales
# --------------------------------------------------------------------------- #
RE_A_CARGO = r'(?:Impuesto a cargo|A cargo)[:\s]+\(?-?([\d,\.]+)\)?'
RE_A_FAVOR = r'(?:Impuesto a favor|A favor)[:\s]+\(?-?([\d,\.]+)\)?'
RE_COMP = r'Compensaciones[:\s]+\(?-?([\d,\.]+)\)?'
RE_CANT_CARGO = r'Cantidad a cargo[:\s]+\(?-?([\d,\.]+)\)?'
RE_CANT_FAVOR = r'Cantidad a favor[:\s]+\(?-?([\d,\.]+)\)?'
RE_CANT_PAGAR = r'Cantidad a pagar[:\s]+\(?-?([\d,\.]+)\)?'


def _parse_origen(body: str) -> Optional[OrigenCompensacion]:
    """Extrae el bloque 'Origen del Saldo a favor / Pago de lo indebido'.

    El bloque se delimita HASTA 'Cantidad a cargo' (no usar 'Concepto de pago'
    como terminador: aparece dentro del bloque sin número y corta de más).
    """
    mm = re.search(
        r'Origen del (Saldo a favor|Pago de lo indebido)(.*?)Cantidad a cargo',
        body, re.S,
    )
    if not mm:
        return None
    tipo_origen = mm.group(1).strip()
    blk = _flat(mm.group(2))
    return OrigenCompensacion(
        tipo_origen=tipo_origen,
        origen_ejercicio=int(v) if (v := _search(r'Ejercicio[:\s]+(\d{4})', blk)) else None,
        saldo_aplicar=num(_search(r'Saldo a aplicar[:\s]+([\d,\.]+)', blk)),
        op_origen=_search(r'N[uú]mero de operaci[oó]n[:\s]+(\d+)', blk),
        monto_original=num(_search(r'Monto del saldo a favor original[:\s]+([\d,\.]+)', blk)),
        remanente_hist=num(_search(
            r'Remanente hist[oó]rico antes de la aplicaci[oó]n[:\s]+([\d,\.]+)', blk)),
        remanente_act=num(_search(
            r'Remanente actualizado antes de la aplicaci[oó]n[:\s]+([\d,\.]+)', blk)),
        fecha_saldo=_search(
            r'Fecha en que se present[oó] la declaraci[oó]n del saldo a favor[:\s]+'
            r'(\d{2}/\d{2}/\d{4})', blk),
    )


def _parse_conceptos_mensual(txt: str) -> list[Concepto]:
    """Divide el cuerpo por el encabezado 'Concepto de pago N:' y parsea cada bloque."""
    headers = re.findall(r'Concepto de pago\s*\d+\s*:\s*([^\n]+)', txt)
    partes = re.split(r'Concepto de pago\s*\d+\s*:', txt)
    conceptos: list[Concepto] = []
    # partes[0] es el preámbulo; partes[1:] alinean con headers.
    for nombre, body in zip(headers, partes[1:]):
        flat = _flat(body)
        origen = _parse_origen(body)
        conceptos.append(Concepto(
            nombre=nombre.strip(),
            a_cargo=num(_search(RE_A_CARGO, flat)),
            a_favor=num(_search(RE_A_FAVOR, flat)),
            cantidad_favor=num(_search(RE_CANT_FAVOR, flat)),
            compensaciones=num(_search(RE_COMP, flat)),
            cantidad_cargo=num(_search(RE_CANT_CARGO, flat)),
            cantidad_pagar=num(_search(RE_CANT_PAGAR, flat)),
            origen=origen,
        ))
    return conceptos


# --------------------------------------------------------------------------- #
# Conceptos en formato REJILLA (acuses de dos columnas / modo de flujo)
# --------------------------------------------------------------------------- #
# Estructura (pdftotext SIN -layout):
#   Concepto de pago 1: A cargo: Cantidad a cargo: Cantidad a pagar:   <- etiquetas
#   Concepto de pago 3: A favor: Cantidad a cargo: Cantidad a pagar:
#   ISR simplificado de confianza. Personas morales 0 0 0             <- nombre + valores
#   IVA simplificado de confianza 24,289 0 0
# La i-ésima línea de datos corresponde a la i-ésima línea de etiquetas; dentro
# de cada línea de datos, los últimos N tokens numéricos mapean a las N etiquetas.
_RE_GRID_LABEL = re.compile(
    r'Concepto de pago\s*(\d+)\s*:(.*)', re.IGNORECASE)
_RE_NUM_TOK = re.compile(r'^\(?-?\$?[\d.,]+\)?$')
_RE_NOMBRE_IMP = re.compile(
    r'^(ISR|IVA|IEPS|ISAN|IABA|Impuesto|Derecho|Cuota)\b', re.IGNORECASE)


def es_formato_grid(raw: str) -> bool:
    """Detecta la rejilla: la etiqueta del concepto trae 'A cargo/favor' inline."""
    return bool(re.search(
        r'Concepto de pago\s*\d+\s*:\s*(?:A cargo|A favor|Impuesto a (?:cargo|favor))',
        raw, re.IGNORECASE))


def _map_labels_a_concepto(nombre: str, labels: list[str], nums: list[str]) -> Concepto:
    c = Concepto(nombre=nombre.strip())
    for lab, val in zip(labels, nums):
        lo = lab.lower()
        v = num(val)
        if 'cantidad a cargo' in lo:
            c.cantidad_cargo = v
        elif 'cantidad a favor' in lo:   # antes de 'a favor': 'cantidad a favor' lo contiene
            c.cantidad_favor = v
        elif 'cantidad a pagar' in lo:
            c.cantidad_pagar = v
        elif 'a favor' in lo:
            c.a_favor = v
        elif 'a cargo' in lo:
            c.a_cargo = v
        elif 'compensacion' in lo:
            c.compensaciones = v
        # 'parte actualizada' y 'recargos' se ignoran (no son base del análisis)
    return c


def _parse_conceptos_grid(raw: str) -> list[Concepto]:
    lineas = [l.strip() for l in raw.split('\n') if l.strip()]

    etiquetas: list[list[str]] = []
    for l in lineas:
        m = _RE_GRID_LABEL.match(l)
        if not m:
            continue
        resto = m.group(2)
        if 'cargo' not in resto.lower() and 'favor' not in resto.lower():
            continue
        labels = [p.strip() for p in resto.split(':') if p.strip()]
        etiquetas.append(labels)

    if not etiquetas:
        return []

    datos: list[tuple[str, list[str]]] = []
    for l in lineas:
        if not _RE_NOMBRE_IMP.match(l):
            continue
        toks = l.split()
        k = 0
        for t in reversed(toks):
            if _RE_NUM_TOK.match(t):
                k += 1
            else:
                break
        if k == 0:
            continue
        datos.append((toks, k))

    conceptos: list[Concepto] = []
    for labels, (toks, k) in zip(etiquetas, datos):
        n = len(labels)
        nums = toks[-n:] if len(toks) >= n else toks[-k:]
        nombre = ' '.join(toks[:len(toks) - len(nums)])
        conceptos.append(_map_labels_a_concepto(nombre, labels, nums))
    return conceptos


# --------------------------------------------------------------------------- #
# Conceptos anuales (formato distinto; un solo concepto ISR)
# --------------------------------------------------------------------------- #
def _parse_conceptos_anual(txt: str, flat: str) -> list[Concepto]:
    nombre = _search(r'Concepto de pago 1[:\s]+([^\n]+)', txt) or 'ISR PERSONAS MORALES'
    return [Concepto(
        nombre=nombre.strip(),
        a_favor=num(_search(r'Impuesto a favor[:\s]+([\d,\.]+)', flat)),
        cantidad_favor=num(_search(RE_CANT_FAVOR, flat)),
        cantidad_cargo=num(_search(r'Cantidad a cargo[:\s]+([\d,\.]+)', flat)),
        cantidad_pagar=num(_search(r'Cantidad a pagar[:\s]+([\d,\.]+)', flat)),
    )]


# --------------------------------------------------------------------------- #
# Entrada principal
# --------------------------------------------------------------------------- #
def _sin_acentos(s: str) -> str:
    return unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode()


def parse_declaracion(texto: str, archivo: str = "",
                      texto_raw: Optional[str] = None) -> Declaracion:
    """Parsea el texto de un acuse en una Declaracion.

    `texto`      : salida de `pdftotext -layout` (metadatos y formato blueprint).
    `texto_raw`  : salida de `pdftotext` sin -layout (para acuses en rejilla);
                   opcional. Si trae formato de rejilla, se usa para los conceptos.
    """
    flat = _flat(texto)
    cab = _sin_acentos(texto[:1000]).upper()
    es_anual = ('DECLARACION ANUAL' in cab or 'ANUAL DE PERSONAS' in cab
                or 'DECLARACION DEL EJERCICIO' in cab)

    rfc = _search(RE_RFC, flat) or ""
    razon = _search(RE_RAZON, texto) or ""
    ejercicio = int(v) if (v := _search(RE_EJERCICIO, flat)) else None

    periodo: Optional[int]
    if es_anual:
        periodo = 0
    else:
        pnombre = _search(RE_PERIODO, flat)
        periodo = MESES.get(pnombre.lower()) if pnombre else None

    tipo = _search(RE_TIPO, texto) or "Normal"
    subtipo = _subtipo(Path(archivo).name, flat)
    if subtipo != 'normal':
        tipo = 'Complementaria'

    if texto_raw and es_formato_grid(texto_raw):
        conceptos = _parse_conceptos_grid(texto_raw)
    elif es_anual:
        conceptos = _parse_conceptos_anual(texto, flat)
    else:
        conceptos = _parse_conceptos_mensual(texto)

    return Declaracion(
        archivo=archivo,
        rfc=rfc,
        razon_social=razon,
        es_anual=es_anual,
        ejercicio=ejercicio,
        periodo=periodo,
        tipo=tipo,
        subtipo=subtipo,
        fecha_presentacion=_fecha_dt(flat),
        operacion=_search(RE_OPERACION, flat),
        version=_search(RE_VERSION, flat),
        conceptos=conceptos,
    )
