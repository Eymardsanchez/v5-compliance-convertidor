"""
bancos/generico.py
──────────────────
Extractor GENÉRICO para bancos NO reconocidos. Hereda toda la maquinaria probada
de ``ExtractorEstandar`` (balance-tracking, clasificación por columna) y lo único
que hace distinto es **auto-detectar el formato de fecha** del PDF antes de
extraer. Así, al subir el estado de cuenta de un banco nuevo, se convierte de
una vez (best-effort) sin escribir código a mano.

Los formatos de fecha soportados en la detección automática cubren lo que usan
los bancos mexicanos:
  · 01/04/2026  01-04-2026  01/04/26           (día/mes/año numérico)
  · 01 ABR 2026   01-ABR-2026   1 ABR           (día MES-3-letras [año])
  · 2026-04-01                                   (ISO)
"""

import re
from collections import Counter
from typing import Optional, Callable

import pdfplumber

from bancos.estandar import BancoEstandar

# Candidatos de patrón de FECHA al inicio de línea (regex origen → compilado).
# El grupo 0 debe capturar la fecha completa; ExtractorEstandar usa m.end() para
# separar el resto de la línea, así que basta con que casen el prefijo de fecha.
_CANDIDATOS_FECHA = [
    r"^(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})",              # 01/04/2026 · 01-04-26
    r"^(\d{4}[/-]\d{1,2}[/-]\d{1,2})",                # 2026-04-01 (ISO)
    r"^(\d{1,2}[ /-][A-Za-z]{3}[ /-]?\d{0,4})",       # 01 ABR 2026 · 01-ABR-26 · 1 ABR
]

# Una línea de transacción típicamente TERMINA con un importe (y a veces saldo).
_MONTO_FIN_RE = re.compile(r"[\d,]+\.\d{2}\s*$")


def detectar_date_re(pdf_path: str, max_pags: int = 4) -> tuple[re.Pattern, str]:
    """Escanea las primeras páginas y elige el patrón de fecha que MÁS líneas de
    transacción explica (líneas que empiezan con fecha y terminan en importe).
    Devuelve (patrón_compilado, regex_origen). Si nada casa, cae al numérico."""
    compilados = [(re.compile(src), src) for src in _CANDIDATOS_FECHA]
    votos: Counter = Counter()
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[:max_pags]:
                texto = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                for raw in texto.splitlines():
                    linea = raw.strip()
                    if not linea or not _MONTO_FIN_RE.search(linea):
                        continue
                    for pat, src in compilados:
                        if pat.match(linea):
                            votos[src] += 1
                            break     # el primero que casa gana para esa línea
    except Exception:
        pass
    if votos:
        mejor_src = votos.most_common(1)[0][0]
        return re.compile(mejor_src), mejor_src
    return compilados[0]     # fallback: numérico d/m/a


class ExtractorGenerico(BancoEstandar):
    """Extractor de último recurso para bancos sin módulo propio."""

    NOMBRE:    str  = "Genérico (auto)"
    COLOR_HEX: str  = "607D8B"
    ES_GENERICO     = True     # marca para que la UI ofrezca generar el parser

    # No exige activar sección: cualquier línea con fecha (que termine en importe)
    # es un movimiento. Así funciona con formatos que no traen "saldo anterior" ni
    # un encabezado de columnas reconocible.
    SECCION_LIBRE = True

    # Activa la sección de movimientos al ver el encabezado de columnas típico.
    SECTION_RE = re.compile(
        r"FECHA.*(DESCRIPCI|CONCEPTO|REFERENC|MOVIMIENTO|DETALLE|OPERACI)",
        re.IGNORECASE)

    # Líneas que traen fecha/importe pero NO son movimientos (encabezados, pies,
    # resúmenes). Reducen el ruido en modo SECCION_LIBRE.
    SKIP_PATTERNS = [
        "SALDO ANTERIOR", "SALDO INICIAL", "SALDO FINAL", "SALDO PROMEDIO",
        "SALDO MINIMO", "SALDO AL", "PERIODO", "FECHA DE CORTE", "CORTE AL",
        "PAGINA", "PÁGINA", "ESTADO DE CUENTA", "RESUMEN", "TOTAL DE",
        "R.F.C", "RFC", "CLABE", "NUMERO DE CUENTA", "NÚMERO DE CUENTA",
        "GAT", "TASA", "COMISIONES", "IPAB", "CADENA ORIGINAL", "SELLO",
        "FOLIO FISCAL", "NO. CLIENTE", "NUMERO DE CLIENTE",
    ]

    # DATE_RE se fija en tiempo de ejecución (auto-detectado por PDF).
    DATE_RE = re.compile(_CANDIDATOS_FECHA[0])

    @classmethod
    def detectar(cls, primera_pagina_texto: str) -> bool:
        # Catch-all: NO se registra en la detección normal; se usa explícitamente
        # como fallback desde detector.detectar_o_generico().
        return False

    def _es_saldo_anterior(self, linea: str) -> Optional[float]:
        u = linea.upper()
        if ("SALDO ANTERIOR" in u or "SALDO INICIAL" in u
                or "BALANCE INICIAL" in u or "SALDO ANT" in u):
            m = re.findall(r"([\d,]+\.\d{2})", linea)
            if m:
                try:
                    return float(m[-1].replace(",", ""))
                except ValueError:
                    return None
        return None

    def extraer(self, pdf_path: str,
                progress_callback: Optional[Callable[[int, int], None]] = None):
        # 1) Auto-detectar el formato de fecha de ESTE PDF.
        self.DATE_RE, self._date_src = detectar_date_re(pdf_path)
        # 2) Si nunca se detectó "saldo anterior" ni el encabezado, como red de
        #    seguridad activamos la sección desde la primera línea con fecha:
        #    se hace forzando SECTION_RE = DATE_RE si el PDF no trae encabezado.
        return super().extraer(pdf_path, progress_callback)
