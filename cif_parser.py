"""
cif_parser.py — Parser de Constancia de Situación Fiscal (CSF/CIF) del SAT.

Lee el PDF oficial del SAT y extrae:
  • RFC
  • Denominación / Razón Social
  • Régimen(es) fiscal(es) — nombre completo y clave SAT

Uso:
    from cif_parser import parsear_cif
    datos = parsear_cif("STE120306P75-constancia.pdf")
    print(datos.rfc, datos.regimenes_codigos)
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pdfplumber


# ── Catálogo de regímenes SAT → código ───────────────────────────────────────
# Claves normalizadas (sin acentos, mayúsculas, sin espacios extras).
# Fuente: Catálogo de Regímenes Fiscales SAT (vigente 2024-2026).

_REGIMEN_A_CLAVE: dict[str, str] = {
    "GENERAL DE LEY PERSONAS MORALES":                                          "601",
    "PERSONAS MORALES CON FINES NO LUCRATIVOS":                                 "603",
    "SUELDOS Y SALARIOS E INGRESOS ASIMILADOS A SALARIOS":                      "605",
    "ARRENDAMIENTO":                                                             "606",
    "REGIMEN DE ENAJENACION O ADQUISICION DE BIENES":                           "607",
    "ENAJENACION O ADQUISICION DE BIENES":                                      "607",
    "DEMAS INGRESOS":                                                            "608",
    "RESIDENTES EN EL EXTRANJERO SIN ESTABLECIMIENTO PERMANENTE EN MEXICO":     "610",
    "INGRESOS POR DIVIDENDOS (SOCIOS Y ACCIONISTAS)":                           "611",
    "INGRESOS POR DIVIDENDOS SOCIOS Y ACCIONISTAS":                             "611",
    "PERSONAS FISICAS CON ACTIVIDADES EMPRESARIALES Y PROFESIONALES":           "612",
    "INGRESOS POR INTERESES":                                                    "614",
    "REGIMEN DE LOS INGRESOS POR OBTENCION DE PREMIOS":                         "615",
    "SIN OBLIGACIONES FISCALES":                                                 "616",
    "SOCIEDADES COOPERATIVAS DE PRODUCCION QUE OPTAN POR DIFERIR SUS INGRESOS": "620",
    "INCORPORACION FISCAL":                                                      "621",
    "ACTIVIDADES AGRICOLAS GANADERAS SILVICOLAS Y PESQUERAS":                   "622",
    "ACTIVIDADES AGRICOLAS, GANADERAS, SILVICOLAS Y PESQUERAS":                 "622",
    "OPCIONAL PARA GRUPOS DE SOCIEDADES":                                        "623",
    "COORDINADOS":                                                               "624",
    "REGIMEN DE LAS ACTIVIDADES EMPRESARIALES CON INGRESOS A TRAVES DE PLATAFORMAS TECNOLOGICAS": "625",
    "PLATAFORMAS TECNOLOGICAS":                                                  "625",
    "REGIMEN SIMPLIFICADO DE CONFIANZA":                                         "626",
    "SIMPLIFICADO DE CONFIANZA":                                                 "626",
    # Regímenes anteriores (CFDI 3.3 aún circulan)
    "PEQUEÑOS CONTRIBUYENTES":                                                   "629",
    "REGIMEN DE PEQUEÑOS CONTRIBUYENTES":                                        "629",
    "REGIMEN DE PEQUEÑOS CONTRIBUYENTES REPECOS":                                "629",
    "ASALARIADOS":                                                               "605",
}

# Patrones de parseo
_RFC_RE   = re.compile(r'\bRFC\b[:\s]+([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})\b', re.IGNORECASE)
_DATE_RE  = re.compile(r'\b\d{2}/\d{2}/\d{4}\b')
_PAGE_RE  = re.compile(r'[Pp]ágina\s+\[\d+\]', re.IGNORECASE)


def _norm(s: str) -> str:
    """Normaliza a mayúsculas sin acentos para lookup en el catálogo."""
    return ''.join(
        c for c in unicodedata.normalize('NFKD', s.upper())
        if not unicodedata.combining(c)
    ).strip()


def _nombre_a_clave(nombre: str) -> str:
    """Retorna la clave SAT del régimen o '' si no se reconoce."""
    n = _norm(nombre)
    if n in _REGIMEN_A_CLAVE:
        return _REGIMEN_A_CLAVE[n]
    # Búsqueda parcial (el PDF puede tener texto recortado)
    for key, code in _REGIMEN_A_CLAVE.items():
        if key in n or n in key:
            return code
    return ""


# ── Dataclass de resultado ────────────────────────────────────────────────────

@dataclass
class DatosCSF:
    rfc:               str        = ""
    razon_social:      str        = ""
    regimenes_nombres: list[str]  = field(default_factory=list)  # nombre completo
    regimenes_codigos: list[str]  = field(default_factory=list)  # clave SAT "626"
    fecha_emision:     str        = ""

    def __str__(self) -> str:
        regs = ", ".join(
            f"{n} ({c})" if c else n
            for n, c in zip(self.regimenes_nombres, self.regimenes_codigos)
        ) or "(sin régimen detectado)"
        return f"{self.rfc} | {self.razon_social} | {regs}"


# ── Nombre de persona física ───────────────────────────────────────────────────

def _nombre_persona_fisica(all_lines: list[str]) -> str:
    """Arma el nombre desde 'Nombre (s)', 'Primer Apellido', 'Segundo Apellido'.

    Las CSF de persona física no traen 'Denominación/Razón Social'; sin esto la
    empresa se registraba con el RFC como nombre (empresas.py). Orden nombre-
    primero para coincidir con cómo aparece en estados de cuenta.
    """
    campos = {"NOMBRE(S)": "", "PRIMERAPELLIDO": "", "SEGUNDOAPELLIDO": ""}
    for i, line in enumerate(all_lines):
        compact = _norm(line).replace(" ", "")
        for k in campos:
            if not campos[k] and compact.startswith(k):
                partes = line.split(":", 1)
                if len(partes) == 2 and partes[1].strip():
                    campos[k] = partes[1].strip()
                elif i + 1 < len(all_lines):
                    campos[k] = all_lines[i + 1].strip()
    full = " ".join(v for v in (campos["NOMBRE(S)"], campos["PRIMERAPELLIDO"],
                                campos["SEGUNDOAPELLIDO"]) if v)
    return " ".join(full.split())


# ── Parser principal ──────────────────────────────────────────────────────────

def parsear_cif(pdf_path: str | Path) -> Optional[DatosCSF]:
    """
    Lee una Constancia de Situación Fiscal en PDF y retorna DatosCSF.
    Retorna None si el archivo no es una CSF válida o no se puede leer.
    """
    try:
        all_lines: list[str] = []
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                txt = page.extract_text(x_tolerance=3, y_tolerance=3) or ""
                all_lines.extend(txt.splitlines())
    except Exception:
        return None

    datos = DatosCSF()

    # ── RFC ───────────────────────────────────────────────────────────────────
    for line in all_lines:
        m = _RFC_RE.search(line)
        if m:
            datos.rfc = m.group(1).upper().strip()
            break

    if not datos.rfc:
        return None   # No es una CSF válida

    # ── Razón Social ──────────────────────────────────────────────────────────
    # El PDF puede producir "Denominación/RazónSocial" (sin espacio) por layout.
    _RAZON_KEYS = ("DENOMINACION/RAZONSOCIAL", "DENOMINACION/RAZON SOCIAL",
                   "DENOMINACION/RAZÓN SOCIAL")
    for i, line in enumerate(all_lines):
        lu_compact = _norm(line).replace(" ", "")
        if any(k in lu_compact for k in _RAZON_KEYS):
            partes = line.split(":", 1)
            if len(partes) == 2 and partes[1].strip():
                datos.razon_social = partes[1].strip()
            elif i + 1 < len(all_lines):
                datos.razon_social = all_lines[i + 1].strip()
            break

    # ── Nombre de persona física (CSF sin "Razón Social") ─────────────────────
    if not datos.razon_social:
        datos.razon_social = _nombre_persona_fisica(all_lines)

    # ── Fecha de emisión ──────────────────────────────────────────────────────
    _DATE_EMISION_RE = re.compile(
        r'A\s+\d{1,2}\s+DE\s+\w+\s+DE\s+\d{4}', re.IGNORECASE
    )
    for i, line in enumerate(all_lines):
        if "LUGAR Y FECHA DE EMISI" in _norm(line):
            # Buscar la primera línea posterior que contenga la fecha
            for extra in all_lines[i + 1: i + 5]:
                if _DATE_EMISION_RE.search(extra):
                    datos.fecha_emision = extra.strip()
                    break
            break

    # ── Regímenes fiscales ────────────────────────────────────────────────────
    in_regimenes  = False
    past_header   = False

    for line in all_lines:
        lu_strip = _norm(line)

        # Marcador de inicio de sección
        if lu_strip in ("REGIMENES:", "REGIMENES") or lu_strip.startswith("REGIMENES:"):
            in_regimenes = True
            past_header  = False
            continue

        if not in_regimenes:
            continue

        # Marcadores de fin de sección
        if (lu_strip.startswith("OBLIGACIONES") or
                lu_strip.startswith("CADENA ORIGINAL") or
                lu_strip.startswith("SELLO DIGITAL") or
                _PAGE_RE.match(line.strip())):
            break

        ls = line.strip()
        if not ls:
            continue

        # Fila de encabezado de tabla ("Régimen   Fecha Inicio   Fecha Fin")
        if not past_header and ("FECHA INICIO" in lu_strip or "FECHA FIN" in lu_strip):
            past_header = True
            continue

        if past_header:
            # Quitar fechas (DD/MM/YYYY) al final
            nombre = _DATE_RE.sub("", ls).strip()
            if not nombre or "PAGINA" in _norm(nombre):
                continue
            # Evitar duplicados
            if nombre not in datos.regimenes_nombres:
                datos.regimenes_nombres.append(nombre)
                datos.regimenes_codigos.append(_nombre_a_clave(nombre))

    return datos
