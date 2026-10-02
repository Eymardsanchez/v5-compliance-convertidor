"""Orquestación reutilizable: leer PDFs -> analizar -> generar XLSX.

Separa la LECTURA (costosa) del ANÁLISIS (barato) para que una interfaz pueda
leer una vez y re-analizar al cambiar de RFC o de saldo inicial. Lo usan tanto
la CLI (`cli.py`) como la app de carga (`app.py`, Streamlit).
"""
from __future__ import annotations

import tempfile
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

from . import dedup as _dedup
from .analyzer_isr import analizar_isr, total_remanente_nominal
from .analyzer_iva import AnalisisIVA, analizar_iva
from .analyzer_iva_acumulado import AnalisisIVAAcumulado, analizar_iva_acumulado
from .config import clasifica_concepto
from .extractor import iter_pdfs, pdf_a_texto, pdf_a_texto_raw
from .models import AcusePago, Declaracion, DetalleIVA, OrigenSaldo
from .parser import parse_declaracion
from .parser_antigua import es_declaracion_antigua, parse_antigua_iva
from .parser_detalle import es_declaracion_detallada, parse_detalle_iva
from .parser_pago import es_acuse_pago, parse_acuse_pago
from .report_xlsx import generar_reporte


# --------------------------------------------------------------------------- #
@dataclass
class Cargados:
    """Resultado de la lectura de PDFs (independiente del RFC)."""
    declaraciones: list[Declaracion] = field(default_factory=list)
    detalles: list[DetalleIVA] = field(default_factory=list)
    pagos: list[AcusePago] = field(default_factory=list)
    no_legibles: list[str] = field(default_factory=list)
    n_pdfs: int = 0

    def rfcs(self) -> list[str]:
        s = {d.rfc for d in self.declaraciones if d.rfc}
        s |= {d.rfc for d in self.detalles if d.rfc}
        s |= {p.rfc for p in self.pagos if p.rfc}
        return sorted(s)


@dataclass
class ResultadoAnalisis:
    """Análisis para un RFC concreto (listo para reporte o UI)."""
    rfc: str
    razon_social: str
    regimen: str
    origenes_isr: list[OrigenSaldo]
    iva: AnalisisIVA
    iva_acumulado: Optional[AnalisisIVAAcumulado]
    n_unicas: int
    n_vigentes: int
    n_anuales: int
    n_pdfs: int
    no_legibles: list[str]
    pagos: list[AcusePago] = field(default_factory=list)

    @property
    def remanente_isr_total(self) -> float:
        return total_remanente_nominal(self.origenes_isr)

    @property
    def total_pagado(self) -> float:
        return round(sum(p.importe_pagado or 0 for p in self.pagos), 2)


# --------------------------------------------------------------------------- #
def expandir_entradas(entradas: Iterable[Path], destino_zip: Optional[Path] = None) -> list[Path]:
    """Normaliza entradas (carpetas, PDFs, ZIPs) a una lista de rutas .pdf.

    Los ZIP se extraen a `destino_zip` (o a un temporal). Recorre carpetas de
    forma recursiva.
    """
    entradas = [Path(e) for e in entradas]
    if destino_zip is None:
        destino_zip = Path(tempfile.mkdtemp(prefix="sat_zip_"))
    pdfs: list[Path] = []
    for e in entradas:
        if e.is_dir():
            pdfs.extend(iter_pdfs(e))
            for z in e.rglob('*.zip'):        # ZIPs anidados en la carpeta
                try:
                    with zipfile.ZipFile(z) as zf:
                        sub = destino_zip / z.stem
                        zf.extractall(sub)
                        pdfs.extend(iter_pdfs(sub))
                except zipfile.BadZipFile:
                    pass
        elif e.suffix.lower() == '.zip':
            try:
                with zipfile.ZipFile(e) as zf:
                    sub = destino_zip / e.stem
                    zf.extractall(sub)
                    pdfs.extend(iter_pdfs(sub))
            except zipfile.BadZipFile:
                pass
        elif e.suffix.lower() == '.pdf':
            pdfs.append(e)
    # dedup de rutas conservando orden
    vistos: set[str] = set()
    unicas: list[Path] = []
    for p in pdfs:
        k = str(p.resolve())
        if k not in vistos:
            vistos.add(k)
            unicas.append(p)
    return unicas


def cargar(pdf_paths: Iterable[Path], progreso=None) -> Cargados:
    """Lee y parsea una lista de rutas PDF.

    `progreso`: callback opcional (i, total, nombre) para barras de progreso.
    """
    pdf_paths = list(pdf_paths)
    c = Cargados(n_pdfs=len(pdf_paths))
    for i, ruta in enumerate(pdf_paths):
        if progreso:
            try:
                progreso(i, len(pdf_paths), Path(ruta).name)
            except Exception:  # noqa - un fallo de UI no debe abortar la carga
                pass
        # Cada PDF se aísla: un archivo que truene NO tumba toda la carga.
        try:
            _cargar_uno(Path(ruta), c)
        except Exception as e:  # noqa
            c.no_legibles.append(f"{Path(ruta).name} (error: {type(e).__name__}: {e})")
    return c


def _cargar_uno(ruta: Path, c: Cargados) -> None:
    """Lee y clasifica UN PDF, agregándolo al acumulador `c`."""
    texto = pdf_a_texto(ruta)
    if not texto.strip():
        c.no_legibles.append(ruta.name)
        return
    raw = pdf_a_texto_raw(ruta)
    if es_acuse_pago(texto):
        pago = parse_acuse_pago(ruta, texto, texto_raw=raw)
        if pago is not None:
            c.pagos.append(pago)
        return
    if es_declaracion_detallada(texto):
        det = parse_detalle_iva(texto, archivo=str(ruta))
        if det is not None:
            c.detalles.append(det)
        return
    if es_declaracion_antigua(texto):
        # Formato antiguo (Régimen General): extrae la determinación de IVA para
        # extender el remanente acumulado hasta 10 años atrás.
        det = parse_antigua_iva(ruta, texto)
        if det is not None:
            c.detalles.append(det)
        return
    c.declaraciones.append(parse_declaracion(texto, archivo=str(ruta), texto_raw=raw))


def detectar_regimen(vigentes: list[Declaracion], detalles: list[DetalleIVA]) -> str:
    nombres = " ".join(cc.nombre.upper() for d in vigentes for cc in d.conceptos)
    if 'SIMPLIFICADO DE CONFIANZA' in nombres or detalles:
        return "RESICO PM (Cap. XII Tít. VII LISR)"
    return "Régimen General (Tít. II LISR)"


def analizar(cargados: Cargados, rfc: Optional[str] = None,
             saldo_inicial_iva: Optional[float] = None) -> ResultadoAnalisis:
    """Analiza los datos cargados para un RFC (o el único presente)."""
    declaraciones = cargados.declaraciones
    detalles = cargados.detalles
    pagos = cargados.pagos
    if rfc:
        declaraciones = [d for d in declaraciones if d.rfc == rfc]
        detalles = [d for d in detalles if d.rfc == rfc]
        pagos = [p for p in pagos if p.rfc == rfc]
    pagos = _pagos_vigentes(pagos)

    unicas = _dedup.deduplicar(declaraciones)
    vig_map = _dedup.vigentes_mensuales(unicas)
    vigentes = [d for d in vig_map.values() if not _dedup.esta_sin_efecto(d)]
    anuales = _dedup.anuales(unicas)

    origenes = analizar_isr(vigentes, anuales)
    iva = analizar_iva(vigentes)
    iva_acum = (analizar_iva_acumulado(detalles, saldo_inicial=saldo_inicial_iva or 0.0)
                if detalles else None)
    if iva_acum is not None and iva_acum.renglones:
        iva.remanente_iva_determinable = True

    razon = next((d.razon_social for d in unicas if d.razon_social), "")
    rfc_final = rfc or (cargados.rfcs()[0] if cargados.rfcs() else "")
    regimen = detectar_regimen(vigentes, detalles)

    return ResultadoAnalisis(
        rfc=rfc_final, razon_social=razon, regimen=regimen,
        origenes_isr=origenes, iva=iva, iva_acumulado=iva_acum,
        n_unicas=len(unicas), n_vigentes=len(vig_map), n_anuales=len(anuales),
        n_pdfs=cargados.n_pdfs, no_legibles=cargados.no_legibles, pagos=pagos,
    )


def _pagos_vigentes(pagos: list[AcusePago]) -> list[AcusePago]:
    """Dedup por operación y vigente por (ejercicio, periodo) = mayor fecha de
    presentación (una complementaria de pago sustituye a la normal)."""
    byop: dict[str, AcusePago] = {}
    for p in pagos:
        byop.setdefault(p.operacion or p.archivo, p)
    grupos: dict[tuple, list[AcusePago]] = {}
    for p in byop.values():
        grupos.setdefault((p.ejercicio, p.periodo), []).append(p)
    vig = [max(g, key=lambda x: x.fecha_presentacion or datetime(1900, 1, 1))
           for g in grupos.values()]
    return sorted(vig, key=lambda x: (x.ejercicio or 0, x.periodo or 0))


def generar_xlsx(res: ResultadoAnalisis, ruta_salida: Path,
                 saldo_inicial_iva: Optional[float] = None) -> Path:
    """Genera el papel de trabajo XLSX a partir de un ResultadoAnalisis."""
    return generar_reporte(
        Path(ruta_salida), razon_social=res.razon_social, rfc=res.rfc,
        regimen=res.regimen, origenes_isr=res.origenes_isr, analisis_iva=res.iva,
        iva_acumulado=res.iva_acumulado, saldo_inicial_iva=saldo_inicial_iva,
        pagos=res.pagos,
    )
