# -*- coding: utf-8 -*-
"""
ocr_pdf.py — OCR de estados de cuenta escaneados (imagen) SIN cambiar los parsers.

Cuando un PDF es una imagen (escaneo sin capa de texto), ningún parser basado en
pdfplumber puede leerlo. Aquí NO se inyecta texto plano: se genera un PDF NUEVO
con una CAPA DE TEXTO invisible posicionada sobre la imagen (searchable PDF). Los
extractores de banco lo abren con pdfplumber y obtienen `extract_text` y
`extract_words` CON COORDENADAS igual que en un PDF nativo — por eso "respeta la
forma del parser": no hay que tocar ni un extractor.

Motor: Tesseract (vía pytesseract) + PyMuPDF (fitz) para rasterizar y ensamblar.
Precisión: 400 DPI en gris y OCR en español. Si Tesseract no está instalado/
bundleado, las funciones degradan a None (el llamador decide qué mensaje mostrar).
"""

from __future__ import annotations

import io
import os
import sys
import tempfile


# ── Localización del binario de Tesseract (instalado o bundleado) ──────────────

def _preparar_tesseract() -> bool:
    """Configura pytesseract y devuelve True si Tesseract es utilizable."""
    try:
        import pytesseract
    except Exception:
        return False
    # 1) Bundleado junto al .exe (PyInstaller) o en rutas típicas de Windows.
    candidatos = []
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.argv[0] or "")))
    if base:
        candidatos.append(os.path.join(base, "tesseract", "tesseract.exe"))
        candidatos.append(os.path.join(base, "tesseract.exe"))
    candidatos += [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidatos:
        if c and os.path.exists(c):
            pytesseract.pytesseract.tesseract_cmd = c
            break
    # 2) Verificar que responde (esté en PATH o ya configurado arriba).
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def _idioma(pytesseract) -> str:
    """Español si está disponible; si no, inglés (fallback)."""
    try:
        langs = set(pytesseract.get_languages(config=""))
        if "spa" in langs:
            return "spa"
    except Exception:
        pass
    return "eng"


def ocr_disponible() -> bool:
    """True si se puede hacer OCR (libs + binario de Tesseract presentes)."""
    try:
        import fitz  # noqa: F401  (PyMuPDF)
        from PIL import Image  # noqa: F401
    except Exception:
        return False
    return _preparar_tesseract()


# ── Detección de PDF-imagen (sin capa de texto) ────────────────────────────────

def pdf_tiene_texto(pdf_path: str, paginas: int = 4) -> bool:
    """True si el PDF tiene texto extraíble en alguna de las primeras `paginas`.
    False = escaneo sin OCR (imagen pura). Ante error de lectura devuelve True
    para no bloquear el flujo normal."""
    try:
        import fitz
        doc = fitz.open(pdf_path)
        try:
            for i in range(min(paginas, doc.page_count)):
                if (doc[i].get_text("text") or "").strip():
                    return True
        finally:
            doc.close()
        return False
    except Exception:
        return True


# ── OCR → PDF con capa de texto ────────────────────────────────────────────────

def ocr_a_pdf_buscable(pdf_path: str, dpi: int = 400, log=None) -> str | None:
    """Genera un PDF con capa de texto (OCR) en un archivo temporal y devuelve su
    ruta, o None si no hay OCR disponible o algo falla. La geometría se conserva:
    el resultado se lee con pdfplumber igual que un PDF nativo.

    El llamador es responsable de borrar el temporal cuando termine."""
    try:
        import fitz
        import pytesseract
        from PIL import Image
    except Exception:
        return None
    if not _preparar_tesseract():
        return None
    lang = _idioma(pytesseract)

    def _log(msg):
        if log:
            try:
                log(msg, "info")
            except Exception:
                pass

    src = out = None
    try:
        src = fitz.open(pdf_path)
        out = fitz.open()
        n = src.page_count
        for i in range(n):
            _log(f"    OCR página {i + 1}/{n}…")
            pix = src[i].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            # image_to_pdf_or_hocr('pdf') = imagen + texto OCR invisible posicionado.
            pdf_pagina = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="pdf")
            out.insert_pdf(fitz.open("pdf", pdf_pagina))
        if out.page_count == 0:
            return None
        fd, tmp = tempfile.mkstemp(suffix="_ocr.pdf")
        os.close(fd)
        out.save(tmp)
        return tmp
    except Exception as e:
        _log(f"    OCR falló: {e}")
        return None
    finally:
        for d in (src, out):
            try:
                if d is not None:
                    d.close()
            except Exception:
                pass


if __name__ == "__main__":
    # Autoprueba mínima: reporta disponibilidad y, si se pasa un PDF, lo OCRea.
    print("OCR disponible:", ocr_disponible())
    if len(sys.argv) > 1:
        p = sys.argv[1]
        print("tiene texto:", pdf_tiene_texto(p))
        r = ocr_a_pdf_buscable(p, log=lambda m, _l="": print(m))
        print("resultado:", r)
        if r:
            print("tiene texto tras OCR:", pdf_tiene_texto(r))
