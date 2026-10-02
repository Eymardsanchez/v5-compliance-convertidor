"""CLI de orquestación. Referencia: blueprint sección 14."""
from __future__ import annotations

import json
import subprocess
import sys

# En consolas legacy de Windows (cp1252) rich/typer fallan al emitir caracteres
# Unicode (spinner braille, bordes de tabla, ⚠). Forzar UTF-8 en la salida evita
# el UnicodeEncodeError sin perder el formato en terminales modernas.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")  # type: ignore[attr-defined]
    except Exception:
        pass
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.progress import Progress, SpinnerColumn, TextColumn
from rich.table import Table

from . import pipeline
from .analyzer_iva_acumulado import etiqueta_mes
from .config import MES_ABR, clasifica_concepto
from .extractor import pdf_a_texto, pdf_a_texto_raw
from .parser import parse_declaracion
from .parser_detalle import parse_detalle_iva

app = typer.Typer(add_completion=False, help="Analizador de saldos a favor ISR/IVA (acuses SAT).")
console = Console()


def _to_jsonable(obj):
    if is_dataclass(obj):
        return {k: _to_jsonable(v) for k, v in asdict(obj).items()}
    if isinstance(obj, list):
        return [_to_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_jsonable(v) for k, v in obj.items()}
    return obj


# --------------------------------------------------------------------------- #
@app.command()
def analizar(
    entrada: Path = typer.Option(..., "--entrada", "-e", help="Carpeta, PDF o ZIP con acuses."),
    salida: Path = typer.Option(..., "--salida", "-s", help="Ruta del .xlsx a generar."),
    rfc: Optional[str] = typer.Option(None, "--rfc", help="Filtra/valida por RFC."),
    recalc: bool = typer.Option(False, "--recalc", help="Cachea fórmulas con LibreOffice."),
    saldo_inicial_iva: Optional[float] = typer.Option(
        None, "--saldo-inicial-iva",
        help="Saldo a favor de IVA inicial (cierre del ejercicio anterior) para el acumulado."),
    json_out: Optional[Path] = typer.Option(None, "--json", help="Exporta el modelo a JSON."),
    verbose: bool = typer.Option(False, "-v", "--verbose", help="Salida detallada."),
):
    """Analiza acuses/declaraciones del SAT y genera el papel de trabajo .xlsx."""
    if not entrada.exists():
        console.print(f"[red]No existe la ruta:[/] {entrada}")
        raise typer.Exit(1)

    pdfs = pipeline.expandir_entradas([entrada])
    with Progress(SpinnerColumn(), TextColumn("{task.description}"), console=console) as pr:
        t = pr.add_task("Extrayendo...", total=None)
        cargados = pipeline.cargar(
            pdfs, progreso=lambda i, n, nom: pr.update(t, description=f"Leyendo {nom}"))

    if not cargados.declaraciones and not cargados.detalles:
        console.print("[red]No se parseó ningún acuse ni declaración detallada legible.[/]")
        raise typer.Exit(1)

    rfcs = cargados.rfcs()
    if not rfc and len(rfcs) > 1:
        console.print(f"[red]Múltiples RFC:[/] {', '.join(rfcs)}. Usa --rfc para filtrar.")
        raise typer.Exit(1)

    res = pipeline.analizar(cargados, rfc=rfc, saldo_inicial_iva=saldo_inicial_iva)
    pipeline.generar_xlsx(res, salida, saldo_inicial_iva=saldo_inicial_iva)

    if recalc:
        _recalcular(salida)

    if json_out:
        payload = {
            'rfc': res.rfc, 'razon_social': res.razon_social, 'regimen': res.regimen,
            'isr': _to_jsonable(res.origenes_isr), 'iva': _to_jsonable(res.iva),
            'iva_acumulado': _to_jsonable(res.iva_acumulado) if res.iva_acumulado else None,
        }
        Path(json_out).write_text(json.dumps(payload, ensure_ascii=False, indent=2),
                                  encoding='utf-8')

    _imprimir_resumen(res, salida)


def _imprimir_resumen(res, salida):
    iva = res.iva
    iva_acum = res.iva_acumulado
    console.print()
    t = Table(title=f"Resumen — {res.rfc} {res.razon_social}", show_lines=False)
    t.add_column("Métrica"); t.add_column("Valor", justify="right")
    t.add_row("PDFs encontrados", str(res.n_pdfs))
    t.add_row("Acuses únicos (por operación)", str(res.n_unicas))
    t.add_row("Periodos mensuales vigentes", str(res.n_vigentes))
    t.add_row("Declaraciones anuales", str(res.n_anuales))
    t.add_row("Orígenes de saldo a favor ISR", str(len(res.origenes_isr)))
    t.add_row("[bold]Remanente nominal ISR total[/]",
              f"[bold]${res.remanente_isr_total:,.2f}[/]")
    t.add_row("Meses IVA con saldo a favor", str(iva.meses_con_saldo_favor))
    if iva_acum and iva_acum.renglones:
        t.add_row("Declaraciones detalladas IVA",
                  f"{len(iva_acum.renglones)} mes(es){' — completo' if iva_acum.completo else ''}")
        t.add_row("Remanente IVA determinable", "Sí (declaraciones detalladas)")
    else:
        t.add_row("Remanente IVA determinable", "No (ver alcance)")
    console.print(t)

    if res.origenes_isr:
        ti = Table(title="ISR — Remanente por origen")
        ti.add_column("Origen"); ti.add_column("Original", justify="right")
        ti.add_column("Aplicado", justify="right"); ti.add_column("Remanente", justify="right")
        ti.add_column("Estatus")
        for o in res.origenes_isr:
            ti.add_row(o.etiqueta, f"${o.monto_original:,.0f}", f"${o.total_aplicado:,.0f}",
                       f"${o.remanente_nominal:,.0f}", o.estatus)
        console.print(ti)

    if iva_acum and iva_acum.renglones:
        ta = Table(title="IVA — Remanente acumulado (declaraciones detalladas)")
        ta.add_column("Periodo"); ta.add_column("Op."); ta.add_column("Fecha present.")
        ta.add_column("Generado", justify="right")
        ta.add_column("Acreditado", justify="right"); ta.add_column("Rem. (desde 0)", justify="right")
        for r in iva_acum.renglones:
            fecha = r.fecha_presentacion.strftime("%d/%m/%Y") if r.fecha_presentacion else "—"
            ta.add_row(etiqueta_mes(r.ejercicio, r.periodo), r.operacion or "—", fecha,
                       f"${r.a_favor_generado:,.0f}", f"${r.acreditamiento:,.0f}",
                       f"${r.remanente_desde_cero:,.0f}")
        console.print(ta)
        console.print(f"[bold]Generado:[/] ${iva_acum.total_generado:,.2f}  "
                      f"[bold]Acreditado:[/] ${iva_acum.total_acreditado:,.2f}  "
                      f"[bold]Neto:[/] ${iva_acum.neto:,.2f}")
        if iva_acum.saldo_inicial_requerido > 0:
            console.print(f"[bold red]⚠ Saldo a favor inicial requerido (del año previo) ≥ "
                          f"${iva_acum.saldo_inicial_requerido:,.2f}[/]")
        if iva_acum.capas_cierre:
            origen = " · ".join(
                f"{c.origen} (op {c.operacion or '—'}): ${c.restante:,.0f}"
                for c in iva_acum.capas_cierre)
            console.print(f"[bold green]Origen del remanente al cierre (PEPS):[/] {origen}")
        console.print(f"[cyan]{iva_acum.nota}[/]")

    if res.pagos:
        tp = Table(title="Pagos — Verificación de declaraciones pagadas")
        tp.add_column("Periodo"); tp.add_column("Tipo")
        tp.add_column("Contrib.", justify="right"); tp.add_column("Actualiz.", justify="right")
        tp.add_column("Recargos", justify="right"); tp.add_column("A pagar", justify="right")
        tp.add_column("Pagado", justify="right"); tp.add_column("Fecha pago"); tp.add_column("Estatus")
        for p in res.pagos:
            per = etiqueta_mes(p.ejercicio, p.periodo) if p.periodo else f"{p.ejercicio}"
            est = "[green]✓ Pagado[/]" if p.pagado else "[red]⚠ Sin comprob.[/]"
            tp.add_row(per, p.tipo[:4], f"${p.contribucion:,.0f}", f"${p.actualizacion:,.0f}",
                       f"${p.recargos:,.0f}", f"${p.importe_a_pagar:,.0f}",
                       f"${(p.importe_pagado or 0):,.0f}", p.fecha_pago or "—", est)
        console.print(tp)
        console.print(f"[bold]Total pagado:[/] ${res.total_pagado:,.2f}  "
                      f"[bold]Recargos:[/] ${sum(p.recargos for p in res.pagos):,.2f}  "
                      f"[bold]Actualización:[/] ${sum(p.actualizacion for p in res.pagos):,.2f}")

    if res.no_legibles:
        console.print(f"[yellow]{len(res.no_legibles)} PDF no legibles:[/] "
                      + ", ".join(res.no_legibles[:5]) + (" ..." if len(res.no_legibles) > 5 else ""))
    if not (iva_acum and iva_acum.renglones):
        console.print(f"[yellow]{iva.texto_alcance}[/]")
    console.print(f"[green]✓ Papel de trabajo:[/] {salida}")


def _recalcular(xlsx: Path) -> None:
    """Cachea valores de fórmulas con LibreOffice headless (best-effort)."""
    import shutil
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        console.print("[yellow]LibreOffice no encontrado; se omite --recalc.[/]")
        return
    try:
        subprocess.run([soffice, '--headless', '--calc', '--convert-to', 'xlsx',
                        '--outdir', str(xlsx.parent), str(xlsx)],
                       capture_output=True, timeout=120)
        console.print("[green]✓ Recalculado con LibreOffice.[/]")
    except Exception as e:
        console.print(f"[yellow]Recalc falló:[/] {e}")


@app.command()
def inspeccionar(pdf: Path = typer.Argument(..., help="Acuse PDF a depurar.")):
    """Imprime el parseo de un solo acuse (depuración)."""
    if not pdf.exists():
        console.print(f"[red]No existe:[/] {pdf}")
        raise typer.Exit(1)
    d = parse_declaracion(pdf_a_texto(pdf), archivo=str(pdf), texto_raw=pdf_a_texto_raw(pdf))
    console.print(f"[bold]Archivo:[/] {d.archivo}")
    console.print(f"RFC: {d.rfc} | {d.razon_social}")
    per = "Anual" if d.es_anual else (MES_ABR[d.periodo] if d.periodo else "?")
    console.print(f"Ejercicio: {d.ejercicio} | Periodo: {per} | Tipo: {d.tipo} "
                  f"({d.subtipo}) | Op: {d.operacion} | Fecha: {d.fecha_presentacion}")
    for c in d.conceptos:
        console.print(f"  • [{clasifica_concepto(c.nombre)}] {c.nombre}")
        console.print(f"      a_cargo={c.a_cargo} a_favor={c.a_favor} "
                      f"comp={c.compensaciones} cant_pagar={c.cantidad_pagar}")
        if c.origen:
            o = c.origen
            console.print(f"      origen[{o.tipo_origen}] aplicar={o.saldo_aplicar} "
                          f"original={o.monto_original} fecha={o.fecha_saldo} "
                          f"rem_act={o.remanente_act}")


def main():
    app()


if __name__ == "__main__":
    sys.exit(app())
