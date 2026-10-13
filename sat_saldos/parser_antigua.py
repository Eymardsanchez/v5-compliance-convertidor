"""Parser del formato ANTIGUO de declaración (pre-RESICO, Régimen General).

'DECLARACIÓN PROVISIONAL O DEFINITIVA DE IMPUESTOS FEDERALES' con determinación
completa de ISR e IVA (formato ~2014-2021). Permite el seguimiento de saldos a
favor de IVA hasta 10 años atrás, alimentando el mismo remanente acumulado.

El formulario es denso, a dos columnas y con etiquetas partidas en varias líneas;
`pdftotext` lo desordena. Se reconstruye con las coordenadas de `pdfplumber`:
cada número se asocia a las palabras de su misma banda horizontal (la etiqueta),
y se localizan los campos por fragmentos distintivos de etiqueta.

Campos de IVA que se extraen (página de determinación de IVA):
  - CANTIDAD A CARGO (IVA a cargo del periodo, antes de acreditamiento)
  - ACREDITAMIENTO DEL SALDO A FAVOR DE PERIODOS ANTERIORES (consumo)
  - SALDO A FAVOR DEL PERIODO (saldo a favor generado)
  - IMPUESTO A CARGO (a cargo final, tras acreditamiento)
Verificación: cantidad_a_cargo − acreditamiento = impuesto_a_cargo.
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Optional

from .config import MESES
from .models import DetalleIVA
from .parser import _sin_acentos


def es_declaracion_antigua(texto: str) -> bool:
    n = _sin_acentos(texto).upper()
    return ("DECLARACION PROVISIONAL O DEFINITIVA DE IMPUESTOS FEDERALES" in n
            and "DETERMINACION DEL IMPUESTO" in n
            and "SIMPLIFICADO DE CONFIANZA" not in n)  # excluye detalladas RESICO


def _tiene_iva(texto: str) -> bool:
    return "VALOR AGREGADO" in _sin_acentos(texto).upper()


# --------------------------------------------------------------------------- #
def _num(s: str) -> Optional[float]:
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _fecha(s: str) -> Optional[datetime]:
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%Y"):
        try:
            return datetime.strptime(s.strip(), fmt)
        except ValueError:
            continue
    return None


def _meta(texto: str) -> dict:
    n = _sin_acentos(texto)
    d: dict = {}
    d["rfc"] = (re.search(r"RFC\s+([A-Z&Ñ0-9]{10,13})", n) or [None, None])[1] \
        if re.search(r"RFC\s+([A-Z&Ñ0-9]{10,13})", n) else None
    m = re.search(r"RFC\s+([A-Z&N0-9]{10,13})", n)
    d["rfc"] = m.group(1) if m else None
    m = re.search(r"PERIODO\s+([A-Za-z]+)", n)
    d["periodo"] = MESES.get(m.group(1).lower()) if m else None
    m = re.search(r"EJERCICIO\s+(\d{4})", n) or re.search(r"\b(20\d{2})\b", n)
    d["ejercicio"] = int(m.group(1)) if m else None
    m = re.search(r"OPERACI[O0]N\s+(\d{6,})", n) or re.search(r"\b(\d{9})\b", n)
    d["operacion"] = m.group(1) if m else None
    m = re.search(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})", n)
    d["fecha_presentacion"] = _fecha(m.group(1)) if m else None
    m = re.search(r"TIPO DE DECLARACI[O0]N\s+([A-Za-z]+)", n)
    d["subtipo"] = "complementaria" if (m and m.group(1).lower().startswith("compl")) else "normal"
    return d


VENTANA_ETIQUETA = 240  # pt a la izquierda del número: aísla su columna


def _numeros_etiquetados(palabras) -> list[tuple[float, float, float, str]]:
    """Para cada número, (valor, x0, top, etiqueta).

    La etiqueta son las palabras no numéricas de la MISMA banda horizontal que
    están inmediatamente a la izquierda del número, dentro de una ventana (para no
    mezclar la columna izquierda con la derecha del formulario a dos columnas).
    """
    out = []
    for w in palabras:
        if not re.fullmatch(r"[\d,]+", w["text"]):
            continue
        val = _num(w["text"])
        if val is None:
            continue
        etq = "".join(
            x["text"] for x in palabras
            if abs(x["top"] - w["top"]) < 4
            and not re.fullmatch(r"[\d,]+", x["text"])
            and (w["x0"] - VENTANA_ETIQUETA) <= x["x0"] < w["x0"])
        out.append((val, w["x0"], w["top"], _sin_acentos(etq).upper().replace(" ", "")))
    return out


def _valor_por_etiqueta(nums, fragmento: str, excluir: tuple = ()) -> Optional[float]:
    """Primer número cuya etiqueta contiene `fragmento` y ninguno de `excluir`."""
    for val, _x, _y, etq in nums:
        if fragmento in etq and not any(e in etq for e in excluir):
            return val
    return None


def parse_antigua_iva(ruta, texto: str) -> Optional[DetalleIVA]:
    """Extrae la determinación de IVA del formato antiguo como DetalleIVA.

    Devuelve None si el documento no trae determinación de IVA. Verifica la
    identidad cantidad_a_cargo − acreditamiento = impuesto_a_cargo.
    """
    if not _tiene_iva(texto):
        return None
    try:
        import io

        import pdfplumber

        from .extractor import pdf_bytes
    except ImportError:
        return None

    meta = _meta(texto)
    nums: list = []
    try:
        # Desde bytes hidratados, no por ruta: evita placeholders de OneDrive.
        with pdfplumber.open(io.BytesIO(pdf_bytes(Path(ruta)))) as pdf:
            for pg in pdf.pages:
                t = pg.extract_text() or ""
                nuf = _sin_acentos(t).upper().replace(" ", "")
                # Solo la página de determinación/resultado de IVA. Dos layouts:
                #  - Régimen General clásico: 'IMPUESTO CAUSADO' / 'SALDO A FAVOR DEL PERIODO'.
                #  - DyP 'Determinación de IVA' (p. ej. 2022): 'SALDO A FAVOR' /
                #    'IMPUESTO A FAVOR' (sin el sufijo 'DEL PERIODO'). Se anclan esos
                #    rótulos propios de la determinación para no incluir la página de
                #    'Detalle del pago' (que sí trae 'A FAVOR' pero no estos).
                if "VALORAGREGADO" in nuf and (
                        "IMPUESTOCAUSADO" in nuf
                        or "SALDOAFAVORDELPERIODO" in nuf
                        or "IMPUESTOAFAVOR" in nuf
                        or "SALDOAFAVOR" in nuf):
                    nums += _numeros_etiquetados(pg.extract_words())
    except Exception:
        return None
    if not nums:
        return None

    # Saldo a favor generado del periodo. Orden de preferencia por layout:
    #  1) 'SALDO A FAVOR DEL PERIODO' (Régimen General clásico),
    #  2) 'SALDO A FAVOR' a secas (DyP 2022) — excluye el renglón de acreditamiento
    #     de periodos anteriores, el remanente IEPS y 'OTRAS CANTIDADES A FAVOR',
    #  3) 'IMPUESTO A FAVOR' como respaldo (mismo valor en la determinación DyP).
    a_favor = (
        _valor_por_etiqueta(nums, "SALDOAFAVORDELPERIODO", excluir=("ORIGINAL",))
        or _valor_por_etiqueta(
            nums, "SALDOAFAVOR",
            excluir=("ORIGINAL", "ANTERIORES", "PERIODOSANT", "IEPS", "REMANENTE", "OTRAS"))
        or _valor_por_etiqueta(nums, "IMPUESTOAFAVOR", excluir=("OTRAS",))
    )
    acreditamiento = _valor_por_etiqueta(nums, "DEPERIODOSANTERIORES") or 0.0
    cantidad_cargo = _valor_por_etiqueta(nums, "CANTIDADACARGO")
    impuesto_cargo = _valor_por_etiqueta(nums, "IMPUESTOACARGO",
                                         excluir=("DIFERENCIA", "OTRAS", "DEPOSITOS"))

    # Verificación de la identidad de la determinación.
    verificado = None
    if cantidad_cargo is not None and impuesto_cargo is not None:
        verificado = abs((cantidad_cargo - acreditamiento) - impuesto_cargo) < 1.0

    return DetalleIVA(
        archivo=str(ruta),
        rfc=meta.get("rfc") or "",
        ejercicio=meta.get("ejercicio"),
        periodo=meta.get("periodo"),
        operacion=meta.get("operacion"),
        fecha_presentacion=meta.get("fecha_presentacion"),
        subtipo=meta.get("subtipo") or "normal",
        a_cargo_periodo=cantidad_cargo,
        a_favor_periodo=a_favor if (a_favor or 0) > 0 else None,
        acreditamiento_anterior=acreditamiento or 0.0,
        impuesto_a_cargo_final=impuesto_cargo,
        cantidad_a_pagar=impuesto_cargo,
    )
