"""Parser del acuse 'INFORMACIÓN REGISTRADA DE PAGOS DE CONTRIBUCIONES FEDERALES'.

Este acuse (con SECCIÓN LÍNEA DE CAPTURA e INFORMACIÓN DEL PAGO RECIBIDO) permite
verificar si la declaración realmente se pagó y desglosar la parte actualizada
(Art. 17-A CFF) y los recargos (Art. 21 CFF) del importe pagado.

La rejilla de conceptos viene desalineada en `pdftotext`, así que se reconstruye
con las coordenadas de `pdfplumber` (valor alineado a la derecha, etiqueta a la
izquierda). Las secciones de línea de captura y pago recibido se leen por regex.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from .config import MESES
from .models import AcusePago, ConceptoPago
from .parser import _sin_acentos


def es_acuse_pago(texto: str) -> bool:
    n = _sin_acentos(texto).upper()
    return ("REGISTRADA DE PAGOS DE CONTRIBUCIONES" in n
            and "LINEA DE CAPTURA" in n)


# --------------------------------------------------------------------------- #
def _num(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    s = s.replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _fecha_hora(s: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None


def _busca(patron: str, texto: str, grupo: int = 1) -> Optional[str]:
    m = re.search(patron, texto, re.IGNORECASE)
    return m.group(grupo).strip() if m else None


def _meta(texto: str) -> dict:
    """Extrae metadatos del encabezado (RFC, periodo, operación, fechas)."""
    n = _sin_acentos(texto)
    d: dict = {}
    d["rfc"] = _busca(r"RFC:\s*([A-Z&Ñ0-9]{10,13})", n)
    d["razon_social"] = _busca(r"raz[o0]n social:\s*(.+)", n)
    d["operacion"] = _busca(r"mero de operaci[o0]n:\s*(\d+)", n)
    ej = _busca(r"Ejercicio:\s*(\d{4})", n)
    d["ejercicio"] = int(ej) if ej else None
    mes = _busca(r"per[i1]odo de la declaraci[o0]n:\s*([A-Za-z]+)", n)
    d["periodo"] = MESES.get(mes.lower()) if mes else None
    tipo = _busca(r"Tipo de declaraci[o0]n:\s*([A-Za-z]+)", n)
    d["tipo"] = tipo or "Normal"
    subtipo = "normal"
    if tipo and tipo.lower().startswith("compl"):
        tc = _busca(r"Tipo de Complementaria:\s*([A-Za-z ]+)", n) or ""
        subtipo = "modificacion" if "odificac" in tc.lower() else "complementaria"
    d["subtipo"] = subtipo
    fp = _busca(r"Fecha y hora de presentaci[o0]n:\s*([0-9/]+\s+[0-9:]+)", n)
    d["fecha_presentacion"] = _fecha_hora(fp) if fp else None
    d["vencimiento"] = _busca(r"Vencimiento Obligaci[o0]n:\s*([0-9/]+)", n)
    return d


def _linea_captura(raw: str) -> dict:
    """Importe total a pagar, línea de captura y vigencia (texto RAW, sin -layout).

    Tolera que pdftotext desordene las palabras del rótulo (p. ej. 'Importe a
    pagar:' / 'total' / '$2,270' en líneas separadas): el importe es el monto en
    '$' dentro del segmento de la línea de captura, antes de 'Vigente hasta'.
    """
    n = _sin_acentos(raw)
    d: dict = {}
    i = n.find("LINEA DE CAPTURA")
    seg = n[i:i + 800] if i >= 0 else n
    # Clave de captura: 5 grupos de 4 caracteres, entre 'Captura:' e 'Importe'.
    mcap = re.search(r"Captura:\s*(.*?)\s*Importe", seg, re.DOTALL)
    zona = mcap.group(1) if mcap else seg
    m = re.search(r"([0-9A-Z]{4}(?:\s+[0-9A-Z]{4}){3,4})", zona)
    d["linea_captura"] = re.sub(r"\s+", "", m.group(1)) if m else None
    # Importe a pagar: el último monto en '$' antes de 'Vigente hasta'.
    seg_imp = seg.split("Vigente hasta")[0]
    montos = re.findall(r"\$\s*([\d,]+(?:\.\d+)?)", seg_imp)
    d["importe_a_pagar"] = _num(montos[-1]) if montos else 0.0
    d["vigente_hasta"] = _busca(r"Vigente hasta:\s*([0-9/]+)", n)
    return d


def _pago_recibido(raw: str) -> dict:
    """Sección 'INFORMACIÓN DEL PAGO RECIBIDO' (texto RAW): prueba de pago."""
    n = _sin_acentos(raw)
    d: dict = {"institucion": None, "fecha_pago": None, "importe_pagado": None,
               "medio_pago": None, "operacion_bancaria": None}
    i = n.find("PAGO RECIBIDO")
    if i < 0:
        return d
    lineas = [ln.strip() for ln in n[i:i + 900].splitlines() if ln.strip()]
    # Estructura fija: [titulo, fila-etiquetas, institucion, "<linea> $importe",
    #                   fila-etiquetas2, "<fecha> <medio> <op> <llave>", ...]
    for k, ln in enumerate(lineas):
        if "credito:" in ln.lower() and k + 2 < len(lineas):
            d["institucion"] = lineas[k + 1]
            mimp = re.search(r"\$?\s*([\d,]+(?:\.\d+)?)\s*$", lineas[k + 2])
            if mimp:
                d["importe_pagado"] = _num(mimp.group(1))
        if "fecha del pago" in ln.lower() and k + 1 < len(lineas):
            vals = lineas[k + 1]
            mf = re.match(r"(\d{2}/\d{2}/\d{4})\s+(\S+)\s+(\d{6,})", vals)
            if mf:
                d["fecha_pago"] = mf.group(1)
                d["medio_pago"] = mf.group(2)
                d["operacion_bancaria"] = mf.group(3)
            else:
                d["fecha_pago"] = _busca(r"(\d{2}/\d{2}/\d{4})", vals)
    return d


# --------------------------------------------------------------------------- #
_ETIQUETAS = {
    "a cargo": "a_cargo",
    "actualizacion": "actualizacion",
    "recargos": "recargos",
    "cantidad a cargo": "cantidad_a_cargo",
    "cantidad a pagar": "cantidad_a_pagar",
}


def _conceptos_por_coordenadas(ruta: Path) -> list[ConceptoPago]:
    """Reconstruye la rejilla de conceptos con coordenadas de pdfplumber.

    Cada renglón: etiqueta a la izquierda (x<220), valor a la derecha (x>500).
    Los renglones 'Concepto de pago N' abren un concepto (nombre en x≈240).
    """
    try:
        import io

        import pdfplumber

        from .extractor import pdf_bytes
    except ImportError:
        return []
    conceptos: list[ConceptoPago] = []
    actual: Optional[ConceptoPago] = None
    try:
        # Desde bytes hidratados, no por ruta: evita placeholders de OneDrive.
        with pdfplumber.open(io.BytesIO(pdf_bytes(Path(ruta)))) as pdf:
            palabras = pdf.pages[0].extract_words(keep_blank_chars=False)
    except Exception:
        return []

    # Agrupa palabras por renglón (coordenada vertical redondeada).
    filas: dict[int, list] = {}
    for w in palabras:
        filas.setdefault(round(w["top"]), []).append(w)

    for top in sorted(filas):
        ws = sorted(filas[top], key=lambda w: w["x0"])
        izq = " ".join(w["text"] for w in ws if w["x0"] < 220).strip()
        der = [w["text"] for w in ws if w["x0"] > 500]
        med = " ".join(w["text"] for w in ws if 220 <= w["x0"] <= 500).strip()
        izq_n = _sin_acentos(izq).lower().rstrip(":").strip()
        valor = _num(der[-1]) if der else None

        if izq_n.startswith("concepto de pago") or izq_n.startswith("impuestos que declara"):
            if med:  # abre concepto con su nombre
                actual = ConceptoPago(nombre=med)
                conceptos.append(actual)
            continue
        if actual is None:
            continue
        campo = _ETIQUETAS.get(izq_n)
        if campo and valor is not None:
            setattr(actual, campo, valor)
    return conceptos


# --------------------------------------------------------------------------- #
def parse_acuse_pago(ruta, texto: str, texto_raw: Optional[str] = None) -> Optional[AcusePago]:
    """Parsea un acuse de pago desde su ruta, texto (-layout) y texto RAW.

    `texto`     : pdftotext -layout (encabezado/metadatos).
    `texto_raw` : pdftotext sin -layout (línea de captura y pago recibido, limpios).
    """
    ruta = Path(ruta)
    raw = texto_raw if texto_raw is not None else texto
    meta = _meta(texto)
    lc = _linea_captura(raw)
    pr = _pago_recibido(raw)
    conceptos = _conceptos_por_coordenadas(ruta)

    actualizacion = round(sum(c.actualizacion for c in conceptos), 2)
    recargos = round(sum(c.recargos for c in conceptos), 2)
    importe = lc.get("importe_a_pagar") or 0.0
    # Contribución anclada al importe autoritativo de la línea de captura, para
    # que contribución + actualización + recargos = importe a pagar exactamente
    # (evita el redondeo del SAT entre 'A cargo' y 'Cantidad a pagar').
    if importe > 0:
        contribucion = round(importe - actualizacion - recargos, 2)
    else:
        contribucion = round(sum(c.a_cargo for c in conceptos), 2)

    return AcusePago(
        archivo=str(ruta),
        rfc=meta.get("rfc") or "",
        razon_social=meta.get("razon_social") or "",
        ejercicio=meta.get("ejercicio"),
        periodo=meta.get("periodo"),
        tipo=meta.get("tipo") or "Normal",
        subtipo=meta.get("subtipo") or "normal",
        operacion=meta.get("operacion"),
        fecha_presentacion=meta.get("fecha_presentacion"),
        vencimiento=meta.get("vencimiento"),
        linea_captura=lc.get("linea_captura"),
        importe_a_pagar=lc.get("importe_a_pagar") or 0.0,
        vigente_hasta=lc.get("vigente_hasta"),
        contribucion=contribucion,
        actualizacion=actualizacion,
        recargos=recargos,
        conceptos=conceptos,
        institucion=pr.get("institucion"),
        fecha_pago=pr.get("fecha_pago"),
        importe_pagado=pr.get("importe_pagado"),
        medio_pago=pr.get("medio_pago"),
        operacion_bancaria=pr.get("operacion_bancaria"),
    )
