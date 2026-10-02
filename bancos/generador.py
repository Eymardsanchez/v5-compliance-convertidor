# -*- coding: utf-8 -*-
"""
bancos/generador.py
───────────────────
Genera el CÓDIGO de un extractor específico para un banco NO reconocido, a partir
de su PDF. Analiza el estado de cuenta (nombre del banco, formato de fecha) y
escribe un módulo ``bancos/<slug>.py`` que hereda de ``BancoEstandar`` — el mismo
patrón probado que usan los bancos existentes — pre-rellenado para que solo haya
que afinarlo.

El módulo generado se recoge solo en el siguiente arranque (cargador dinámico en
detector.py), así que el banco queda "aprendido".
"""
from __future__ import annotations

import os
import re
import unicodedata

import pdfplumber

from bancos.generico import detectar_date_re, ExtractorGenerico

CARPETA_BANCOS = os.path.dirname(__file__)


def _sin_acentos(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def _nombre_banco(pdf_path: str) -> str:
    """Extrae el nombre del banco por el ancla 'Institución de Banca Múltiple'
    (lo que está a la izquierda de 'INSTITUC…'). Si no lo halla, usa la primera
    línea prominente de la página 1."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:3]:
                words = page.extract_words() or []
                for w in words:
                    if not _sin_acentos(w["text"]).upper().startswith("INSTITUC"):
                        continue
                    yc = (w["top"] + w["bottom"]) / 2
                    linea = [x for x in words if abs((x["top"] + x["bottom"]) / 2 - yc) < 3]
                    ln = " ".join(_sin_acentos(x["text"]).upper() for x in linea)
                    if "BANCA" not in ln or "MULTIPLE" not in ln:
                        continue
                    izq = sorted([x for x in linea if x["x0"] < w["x0"] - 1],
                                 key=lambda x: x["x0"])
                    if izq:
                        raw = " ".join(x["text"] for x in izq)
                        raw = re.sub(r"[\s,\.]*S\.?\s*A\.?[\s,\.]*$", "", raw).strip(" ,.")
                        if raw:
                            return raw
            # Fallback: buscar una línea con pinta de NOMBRE de banco en las
            # 2 primeras páginas (evitando números de página, fechas y basura).
            def _es_nombre(ln: str) -> bool:
                u = _sin_acentos(ln).upper().strip()
                if re.match(r"^(PAGINA|PAGE|HOJA)\b", u):
                    return False
                if re.match(r"^[\d\s/.\-,:$]+$", ln):       # solo números/fechas
                    return False
                if len(ln.strip()) < 5:
                    return False
                return sum(c.isalpha() for c in ln) >= 4

            _PISTAS = ("BANCO", "S.A.", "S.A.B", "CITIBANAMEX", "BANAMEX",
                       "SCOTIA", "AFIRME", "MIFEL", "MULTIVA", "INBURSA",
                       "COMPARTAMOS", "AZTECA", "BAJIO", "REGIONAL", "INSTITUC")
            with pdfplumber.open(pdf_path) as pdf2:
                candidatos = []
                for page in pdf2.pages[:2]:
                    for ln in (page.extract_text() or "").splitlines():
                        ln = ln.strip()
                        if _es_nombre(ln):
                            candidatos.append(ln)
            # 1º: línea con indicio bancario claro
            for ln in candidatos:
                if any(p in _sin_acentos(ln).upper() for p in _PISTAS):
                    return ln[:40]
            # 2º: la primera línea "válida"
            if candidatos:
                return candidatos[0][:40]
    except Exception:
        pass
    return "Banco Nuevo"


def _slug(nombre: str) -> str:
    base = _sin_acentos(nombre).lower()
    base = re.sub(r"[^a-z0-9]+", "_", base).strip("_")
    partes = [p for p in base.split("_") if p and p not in
              ("banco", "sa", "de", "del", "la", "el", "grupo", "financiero")]
    return "_".join(partes[:2]) or "banco_nuevo"


def _clase(slug: str) -> str:
    return "".join(p.capitalize() for p in slug.split("_")) + "Extractor"


def _keyword(nombre: str) -> str:
    """Token distintivo para detectar() — la primera palabra significativa."""
    for p in _sin_acentos(nombre).upper().split():
        if len(p) >= 4 and p not in ("BANCO", "GRUPO"):
            return p
    return _sin_acentos(nombre).upper().strip()[:8]


def analizar(pdf_path: str) -> dict:
    nombre = _nombre_banco(pdf_path)
    _pat, date_src = detectar_date_re(pdf_path)
    slug = _slug(nombre)
    return {
        "nombre":   nombre,
        "slug":     slug,
        "clase":    _clase(slug),
        "keyword":  _keyword(nombre),
        "date_src": date_src,
    }


_PLANTILLA = '''# -*- coding: utf-8 -*-
"""
Extractor AUTO-GENERADO para {nombre}.
Generado a partir de: {pdf_base}
Formato de fecha detectado: {date_src}

⚠ REVISA Y AFINA: hereda de BancoEstandar (patrón tabular con balance-tracking).
Si algún movimiento no sale bien, ajusta DATE_RE / SKIP_PATTERNS o sobreescribe
métodos como en los otros bancos de /bancos/ (ver banbajio.py como ejemplo).
"""
import re

from bancos.estandar import BancoEstandar


class {clase}(BancoEstandar):
    NOMBRE:    str = "{nombre}"
    COLOR_HEX: str = "607D8B"

    # No exige activar sección: toda línea con fecha que termine en importe es
    # movimiento. Si el banco trae "SALDO ANTERIOR"/encabezado, puedes ponerlo en
    # False y definir SECTION_RE para más precisión.
    SECCION_LIBRE: bool = True

    DATE_RE = re.compile(r"{date_src}")

    SKIP_PATTERNS = [
        "SALDO ANTERIOR", "SALDO INICIAL", "SALDO FINAL", "SALDO PROMEDIO",
        "PERIODO", "FECHA DE CORTE", "PAGINA", "ESTADO DE CUENTA", "RESUMEN",
        "R.F.C", "RFC", "CLABE", "NUMERO DE CUENTA", "GAT", "TASA",
        "COMISIONES", "IPAB", "CADENA ORIGINAL", "SELLO", "FOLIO FISCAL",
    ]

    @classmethod
    def detectar(cls, primera_pagina_texto: str) -> bool:
        return "{keyword}" in primera_pagina_texto.upper()
'''


def generar_modulo(pdf_path: str, carpeta: str | None = None) -> tuple[str, dict]:
    """Analiza el PDF y ESCRIBE bancos/<slug>.py. Devuelve (ruta, info).
    Si el archivo ya existe, agrega un sufijo _2, _3… para no pisar nada."""
    info = analizar(pdf_path)
    carpeta = carpeta or CARPETA_BANCOS
    slug = info["slug"]
    ruta = os.path.join(carpeta, f"{slug}.py")
    n = 2
    while os.path.exists(ruta):
        ruta = os.path.join(carpeta, f"{slug}_{n}.py")
        n += 1
    codigo = _PLANTILLA.format(
        nombre=info["nombre"].replace('"', "'"),
        pdf_base=os.path.basename(pdf_path).replace('"', "'"),
        date_src=info["date_src"],
        clase=info["clase"],
        keyword=info["keyword"],
    )
    with open(ruta, "w", encoding="utf-8") as fh:
        fh.write(codigo)
    info["ruta"] = ruta
    return ruta, info
