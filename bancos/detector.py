"""
Detecta automáticamente qué banco corresponde a un PDF y retorna el extractor correcto.
Para agregar un banco nuevo: crear su módulo en /bancos/ y registrarlo en EXTRACTORES
y en _IBM_MAP.

Estrategia de detección (dos fases):
  1ª PRIORIDAD — Ancla "Institución de Banca Múltiple":
     En el pie/encabezado de cada página de un estado de cuenta mexicano siempre
     aparece la leyenda:
         <NOMBRE BANCO>  S.A.,  Institución de Banca Múltiple  <GRUPO> …
     Se busca la palabra pivote "INSTITUC…" y se toma todo lo que está a su
     IZQUIERDA en la misma línea: eso es el nombre oficial del banco.
     Esta estrategia es inmune a falsos positivos por SPEIs/transferencias a
     otros bancos que aparezcan en el cuerpo de las transacciones.

  2ª PRIORIDAD — Fallback por keywords en texto completo:
     Si la frase IBM no se encontró (PDF escaneado, tipo de fuente incompatible,
     Inbursa que usa su propia leyenda, HSBC con fuentes Type3), se corre cada
     extractor.detectar() en el orden de EXTRACTORES.
"""

import threading
import pdfplumber
from bancos.base import BancoExtractor

# ── Lock global compartido ────────────────────────────────────────────────────
# pdfminer mantiene caché de fuentes a nivel de proceso. Serializar TODA apertura
# de PDF (detección y extracción) evita corrupción cuando hay múltiples hilos.
# Se usa RLock (reentrant) para que el mismo hilo pueda adquirirlo varias veces
# sin deadlock (p.ej. cuando _detect() adquiere el lock y llama a detectar_banco
# que también lo adquiere internamente).
PDF_LOCK = threading.RLock()

from bancos.santander import SantanderExtractor
from bancos.bbva      import BBVAExtractor
from bancos.banregio  import BanregioExtractor
from bancos.inbursa   import InbursaExtractor
from bancos.banorte   import BanorteExtractor
from bancos.banbajio  import BanbajioExtractor
from bancos.hsbc      import HSBCExtractor
from bancos.banamex   import BanamexExtractor
from bancos.konfio    import KonfioExtractor

# ── Registro de extractores (fallback) ───────────────────────────────────────
# Sólo se usa si la estrategia IBM no logra identificar el banco.
# ORDEN: más específicos primero para evitar falsos positivos.
# Konfio va primero: es SOFOM (sin leyenda IBM) y usa "KONFIO"/"RED AMIGO DAL"
# que son tokens muy específicos — no confunden con ningún banco tradicional.
EXTRACTORES: list[type[BancoExtractor]] = [
    KonfioExtractor,     # "KONFIO" / "RED AMIGO DAL" — SOFOM TC, sin IBM
    InbursaExtractor,    # "CLIENTE INBURSA" + "BALANCE INICIAL" — muy específico
    BanregioExtractor,   # "BANREGIO" / "BANCO REGIONAL"
    BanorteExtractor,    # "BANCO MERCANTIL DEL NORTE" / "ENLACE NEGOCIOS/GLOBAL"
    BanbajioExtractor,   # "BANBAJIO" / "BANCO DEL BAJIO"
    SantanderExtractor,  # "BANCO SANTANDER" / "SANTANDER MEXICO"
    BBVAExtractor,       # "BBVA" — antes que Banamex para evitar falso positivo
    BanamexExtractor,    # "BANAMEX" / "CITIBANAMEX" — sólo busca en encabezado
    HSBCExtractor,       # (cid: en texto crudo — al final
]

# ── Mapa IBM: prefijos del nombre oficial → extractor ─────────────────────────
# Se compara contra el texto normalizado (mayúsculas, sin acentos) que aparece
# a la izquierda de "INSTITUC…" en el pie de página.
# Ordenado de más específico a más general para evitar coincidencias parciales.
_IBM_MAP: list[tuple[str, type[BancoExtractor]]] = [
    # Banamex  (antes que variante genérica "BANCO NACIONAL")
    ("BANCO NACIONAL DE MEXICO",   BanamexExtractor),
    ("CITIBANAMEX",                BanamexExtractor),
    ("BANAMEX",                    BanamexExtractor),
    # Banorte  (razón social completa — evita confusión con "BANCO" genérico)
    ("BANCO MERCANTIL DEL NORTE",  BanorteExtractor),
    ("BANORTE",                    BanorteExtractor),
    # Banregio
    ("BANCO REGIONAL DE MONTERREY", BanregioExtractor),
    ("BANCO REGIONAL",             BanregioExtractor),
    ("BANREGIO",                   BanregioExtractor),
    # Banbajío
    ("BANCO DEL BAJIO",            BanbajioExtractor),
    ("BANBAJIO",                   BanbajioExtractor),
    # Santander
    ("BANCO SANTANDER MEXICO",     SantanderExtractor),
    ("BANCO SANTANDER",            SantanderExtractor),
    ("SANTANDER",                  SantanderExtractor),
    # BBVA  (razón social: "BBVA MEXICO, S.A., IBM…")
    ("BBVA",                       BBVAExtractor),
    # HSBC
    ("HSBC MEXICO",                HSBCExtractor),
    ("HSBC",                       HSBCExtractor),
    # Inbursa  (algunos formatos usan IBM aunque Inbursa también tiene su leyenda propia)
    ("BANCO INBURSA",              InbursaExtractor),
    ("INBURSA",                    InbursaExtractor),
]


# ── Helpers estrategia IBM ────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Normaliza a mayúsculas y reemplaza vocales acentuadas y Ñ."""
    tr = str.maketrans(
        "ÁÉÍÓÚÜÑáéíóúüñ",
        "AEIOUUNaeiouun"
    )
    return s.upper().translate(tr)


def _detectar_por_ibm(
    page_words: list[list[dict]],
) -> type[BancoExtractor] | None:
    """
    Recorre las palabras extraídas de cada página buscando el pivote IBM.
    Retorna la CLASE del extractor (no una instancia) o None si no lo encuentra.

    Algoritmo:
      1. Buscar cualquier palabra cuyo texto normalizado empiece con "INSTITUC".
      2. Verificar que en la misma línea (y ± 3 pt) aparezcan "BANCA" y "MULTIPLE"
         (descarta coincidencias espurias con otras palabras que empiecen con "INSTITUC").
      3. Recolectar todas las palabras de esa línea cuya x0 sea MENOR que la del
         pivote (= el nombre del banco) y unirlas en un string normalizado.
      4. Buscar ese string en _IBM_MAP usando startswith para tolerancia a sufijos
         como "S.A.", comas, abreviaciones de grupo financiero, etc.
    """
    for words in page_words:
        if not words:
            continue

        for w in words:
            if not _norm(w["text"]).startswith("INSTITUC"):
                continue

            # Palabras en la misma línea (tolerancia ±3 pt en centro vertical)
            y_center = (w["top"] + w["bottom"]) / 2
            same_line = [
                x for x in words
                if abs((x["top"] + x["bottom"]) / 2 - y_center) < 3
            ]
            line_norm = " ".join(_norm(x["text"]) for x in same_line)

            # Confirmar que es realmente la leyenda IBM
            if "BANCA" not in line_norm or "MULTIPLE" not in line_norm:
                continue

            # Palabras a la IZQUIERDA del pivote (x0 < pivote_x0 con margen de 1 pt)
            pivot_x = w["x0"]
            left_words = sorted(
                [x for x in same_line if x["x0"] < pivot_x - 1],
                key=lambda x: x["x0"],
            )
            if not left_words:
                continue

            bank_raw = " ".join(_norm(x["text"]) for x in left_words)
            # Limpiar comas, puntos, "S.A." y variantes al final
            # (p. ej. "BBVA MEXICO, S.A.," → "BBVA MEXICO")
            import re as _re
            bank_clean = _re.sub(
                r"[\s,\.]*S\.?\s*A\.?[\s,\.]*$",
                "",
                bank_raw,
            ).strip(" ,.")

            # Buscar en el mapa (el más específico primero)
            for prefix, cls in _IBM_MAP:
                if bank_clean.startswith(prefix):
                    return cls

            # Mapa no reconoce el banco — seguir con fallback
            # (No retornamos None aquí; puede haber más páginas por revisar)

    return None


# ── API pública ───────────────────────────────────────────────────────────────

def detectar_banco(pdf_path: str) -> BancoExtractor | None:
    """
    Lee las primeras páginas del PDF y retorna una instancia del extractor correcto.
    Retorna None si no se reconoce el banco.
    Serializado con PDF_LOCK para evitar corrupción del caché global de pdfminer.
    """
    try:
        with PDF_LOCK:
            with pdfplumber.open(pdf_path) as pdf:
                # 4 páginas: BBVA tiene la página 1 como imagen y el encabezado
                # real en pág. 2; 4 páginas garantizan cubrir ese caso.
                sample_pages = min(4, len(pdf.pages))
                page_words = []
                page_texts = []
                for i in range(sample_pages):
                    page = pdf.pages[i]
                    page_words.append(page.extract_words() or [])
                    page_texts.append(page.extract_text() or "")
    except Exception:
        return None

    # ── Estrategia 1: ancla "Institución de Banca Múltiple" ──────────────────
    cls = _detectar_por_ibm(page_words)
    if cls is not None:
        return cls()

    # ── Estrategia 2: fallback por keywords en texto ──────────────────────────
    texto_upper = "\n".join(page_texts).upper()
    for cls in EXTRACTORES:
        if cls.detectar(texto_upper):
            return cls()

    return None


def banco_nombre(pdf_path: str) -> str:
    """Retorna el nombre del banco o 'Desconocido' si no se reconoce."""
    extractor = detectar_banco(pdf_path)
    return extractor.NOMBRE if extractor else "Desconocido"


# ── Cargador dinámico de extractores generados ────────────────────────────────
# Recoge automáticamente cualquier módulo bancos/<x>.py que declare una subclase
# de BancoExtractor (p.ej. los generados por bancos/generador.py) y la agrega al
# registro. Así un banco "aprendido" queda disponible sin editar este archivo.
_MODULOS_BASE = {"base", "estandar", "detector", "generico", "generador", "__init__"}


def _cargar_extractores_dinamicos() -> None:
    import os
    import importlib
    import inspect
    carpeta = os.path.dirname(__file__)
    ya = set(EXTRACTORES)
    for fn in sorted(os.listdir(carpeta)):
        if not fn.endswith(".py") or fn[:-3] in _MODULOS_BASE:
            continue
        mod_name = f"bancos.{fn[:-3]}"
        try:
            mod = importlib.import_module(mod_name)
        except Exception:
            continue
        for _, obj in inspect.getmembers(mod, inspect.isclass):
            if (issubclass(obj, BancoExtractor) and obj is not BancoExtractor
                    and getattr(obj, "__module__", "") == mod_name
                    and obj not in ya):
                EXTRACTORES.append(obj)
                ya.add(obj)


try:
    _cargar_extractores_dinamicos()
except Exception:
    pass


def detectar_o_generico(pdf_path: str):
    """Como detectar_banco, pero si NINGÚN banco casa devuelve el ExtractorGenerico
    (best-effort) en vez de None. El genérico trae ``ES_GENERICO = True`` para que
    la UI ofrezca generar el parser específico."""
    ext = detectar_banco(pdf_path)
    if ext is not None:
        return ext
    from bancos.generico import ExtractorGenerico
    return ExtractorGenerico()
