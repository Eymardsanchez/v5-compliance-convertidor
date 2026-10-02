"""
polizas_coi.py — Motor de generación de Pólizas Aspel COI

Convierte movimientos bancarios conciliados en pólizas contables
con el layout de importación de Aspel COI.

Tipos generados:
  Ig (Ingreso) — depósito bancario: cobro a cliente
  Eg (Egreso)  — retiro bancario:   pago a proveedor / nómina
  Dr (Diario)  — movimiento sin clasificar

Layout por póliza (columnas A-J):
  Fila header:  A=TipoPol, B=Num, C=Concepto, D=Dia
  Fila partida: B=Cuenta, C=Depto, D=Concepto, E=TC, F=Debe, G=Haber, H=CC, I=Proy
  Fila CFDI:    C=INICIO_CFDI → C=FechaCFDI,D=Serie,E=Folio,F=RFCEmi,G=RFCRec,
                                 H=Monto,I=UUID,J=Tipo → C=FIN_CFDI
  Fila cierre:  B=FIN_PARTIDAS
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import TYPE_CHECKING, Optional

if TYPE_CHECKING:
    from conciliacion import MovConciliado
    from cfdi_parser  import CFDI
    from catalogo_terceros import Tercero

# ── Cuenta contable cuando no se encuentra el tercero ─────────────────────────
_CUENTA_NO_DEFINIDA = "XXXX-000-000"


# ── Dataclasses ───────────────────────────────────────────────────────────────

@dataclass
class PartidaPoliza:
    cuenta:       str
    concepto:     str
    debe:         Optional[float]
    haber:        Optional[float]
    tipo_cambio:  float         = 1.0
    depto:        Optional[int] = None
    centro_costo: Optional[int] = None
    proyecto:     Optional[int] = None
    cfdi:         Optional["CFDI"] = field(default=None, repr=False)


@dataclass
class Poliza:
    tipo:     str               # "Ig" / "Eg" / "Dr"
    numero:   int
    concepto: str               # máx 120 chars
    dia:      int               # día del mes
    partidas: list[PartidaPoliza] = field(default_factory=list)

    @property
    def total_debe(self) -> float:
        return round(sum(p.debe or 0 for p in self.partidas), 2)

    @property
    def total_haber(self) -> float:
        return round(sum(p.haber or 0 for p in self.partidas), 2)

    @property
    def cuadrada(self) -> bool:
        return abs(self.total_debe - self.total_haber) < 0.01


# ── Helpers ───────────────────────────────────────────────────────────────────

def _dia(fecha_str: str) -> int:
    """Extrae el día del mes de una fecha en formato DD/MM/YYYY o YYYY-MM-DD."""
    s = fecha_str.strip()
    try:
        if "/" in s:
            return int(s.split("/")[0])
        if "-" in s:
            return int(s.split("-")[2][:2])
    except (ValueError, IndexError):
        pass
    return 1


def _truncar(texto: str, maximo: int = 120) -> str:
    return texto[:maximo].strip()


def _cuenta_tercero(
    cfdi:      "CFDI",
    es_emisor: bool,
    terceros:  list["Tercero"],
    default:   str,
) -> str:
    """Retorna la cuenta_contable del tercero o el default si no se encuentra."""
    from catalogo_terceros import buscar_tercero
    rfc    = cfdi.rfc_emisor   if es_emisor else cfdi.rfc_receptor
    nombre = cfdi.nombre_emisor if es_emisor else cfdi.nombre_receptor
    t = buscar_tercero(terceros, rfc=rfc, nombre=nombre)
    if t and t.cuenta_contable:
        return t.cuenta_contable
    return default or _CUENTA_NO_DEFINIDA


# ── Generador de pólizas ──────────────────────────────────────────────────────

def generar_polizas(
    conciliados:             list["MovConciliado"],
    rfc_empresa:             str,
    num_inicio:              int,
    cuenta_banco:            str,
    terceros:                list["Tercero"],
    cuenta_default_prov:     str  = "",
    cuenta_default_cli:      str  = "",
    cuenta_nomina:           str  = "",
    incluir_sin_cfdi:        bool = False,
) -> list[Poliza]:
    """
    Genera una póliza por cada movimiento conciliado.

    conciliados          — lista de MovConciliado del motor de conciliación
    rfc_empresa          — RFC para identificar lado empresa en CFDIs
    num_inicio           — número de la primera póliza
    cuenta_banco         — cuenta contable del banco (p.ej. "1110-001-000")
    terceros             — catálogo de clientes/proveedores ya cargado
    cuenta_default_prov  — cuenta fallback cuando no se localiza proveedor
    cuenta_default_cli   — cuenta fallback cuando no se localiza cliente
    cuenta_nomina        — cuenta de gasto de nómina (p.ej. "6100-001-000")
    incluir_sin_cfdi     — si True, genera póliza Dr para movimientos sin CFDI
    """
    polizas: list[Poliza] = []
    num = num_inicio
    rfc_u = rfc_empresa.strip().upper()

    for mc in conciliados:
        cfdi  = mc.cfdi
        banco = mc.banco

        # Movimientos sin CFDI cruzado
        if cfdi is None:
            if not incluir_sin_cfdi:
                continue
            monto = banco.monto
            concepto = _truncar(f"Sin CFDI — {banco.descripcion}")
            tipo_mov  = "Ig" if banco.deposito else "Eg"
            cta_contra = _CUENTA_NO_DEFINIDA
            pol = Poliza(tipo=tipo_mov, numero=num,
                         concepto=concepto, dia=_dia(banco.fecha))
            if banco.deposito:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cuenta_banco, concepto=concepto,
                    debe=monto, haber=None))
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_contra, concepto=concepto,
                    debe=None, haber=monto))
            else:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_contra, concepto=concepto,
                    debe=monto, haber=None))
                pol.partidas.append(PartidaPoliza(
                    cuenta=cuenta_banco, concepto=concepto,
                    debe=None, haber=monto))
            polizas.append(pol)
            num += 1
            continue

        dia    = _dia(banco.fecha)
        todos_cfdis = [cfdi] + list(mc.cfdi_suma)

        # ── EGRESO: retiro bancario (pago a proveedor, nómina, complemento) ──
        if mc.tipo == "recibido":
            monto_total = banco.retiro or banco.monto
            tipo_pol    = "Eg"

            # Nómina (tipo N)
            if cfdi.tipo == "N":
                cta_nom  = cuenta_nomina or cuenta_default_prov or _CUENTA_NO_DEFINIDA
                concepto = _truncar(f"Nómina — {cfdi.nombre_receptor or cfdi.rfc_receptor}")
                pol = Poliza(tipo=tipo_pol, numero=num,
                             concepto=concepto, dia=dia)
                for c in todos_cfdis:
                    monto_c = c.total or monto_total / len(todos_cfdis)
                    pol.partidas.append(PartidaPoliza(
                        cuenta=cta_nom,
                        concepto=_truncar(f"Nómina {c.rfc_receptor}"),
                        debe=round(monto_c, 2), haber=None,
                        cfdi=c,
                    ))
                pol.partidas.append(PartidaPoliza(
                    cuenta=cuenta_banco,
                    concepto=concepto,
                    debe=None, haber=monto_total,
                ))

            # Complemento de pago (tipo P) — registra el pago real de la deuda
            elif cfdi.tipo == "P":
                nombre_prov = cfdi.nombre_emisor or cfdi.rfc_emisor
                concepto = _truncar(f"Complemento de pago — {nombre_prov}")
                pol = Poliza(tipo=tipo_pol, numero=num,
                             concepto=concepto, dia=dia)
                cta_prov = _cuenta_tercero(cfdi, True, terceros, cuenta_default_prov)
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_prov,
                    concepto=_truncar(f"Pago compl. {cfdi.uuid[:8]}"),
                    debe=monto_total, haber=None,
                    cfdi=cfdi,
                ))
                pol.partidas.append(PartidaPoliza(
                    cuenta=cuenta_banco,
                    concepto=concepto,
                    debe=None, haber=monto_total,
                ))

            # Factura de proveedor (tipo I/E)
            else:
                nombre_prov = (cfdi.nombre_emisor or cfdi.rfc_emisor)
                concepto = _truncar(f"Pago — {nombre_prov}")
                pol = Poliza(tipo=tipo_pol, numero=num,
                             concepto=concepto, dia=dia)
                monto_acumulado = 0.0
                for c in todos_cfdis:
                    monto_c = c.total or 0.0
                    monto_acumulado += monto_c
                    cta_prov = _cuenta_tercero(c, True, terceros, cuenta_default_prov)
                    conc_p   = _truncar(f"Pago {c.nombre_emisor or c.rfc_emisor}, {c.folio or c.uuid[:8]}")
                    pol.partidas.append(PartidaPoliza(
                        cuenta=cta_prov,
                        concepto=conc_p,
                        debe=round(monto_c, 2), haber=None,
                        cfdi=c,
                    ))
                # Si la suma de CFDIs no cuadra con el movimiento bancario, ajustar
                monto_banco_real = monto_total
                if abs(monto_acumulado - monto_banco_real) > 0.01 and monto_banco_real:
                    monto_acumulado = monto_banco_real
                pol.partidas.append(PartidaPoliza(
                    cuenta=cuenta_banco,
                    concepto=concepto,
                    debe=None, haber=round(monto_acumulado, 2),
                ))

        # ── INGRESO: depósito bancario (cobro a cliente) ──────────────────────
        elif mc.tipo == "emitido":
            monto_total = banco.deposito or banco.monto
            tipo_pol    = "Ig"
            nombre_cli  = (cfdi.nombre_receptor or cfdi.rfc_receptor)
            concepto    = _truncar(f"Cobro — {nombre_cli}")
            pol = Poliza(tipo=tipo_pol, numero=num,
                         concepto=concepto, dia=dia)

            pol.partidas.append(PartidaPoliza(
                cuenta=cuenta_banco,
                concepto=concepto,
                debe=monto_total, haber=None,
            ))
            for c in todos_cfdis:
                monto_c = c.total or 0.0
                cta_cli  = _cuenta_tercero(c, False, terceros, cuenta_default_cli)
                conc_c   = _truncar(f"Cobro {c.nombre_receptor or c.rfc_receptor}, {c.folio or c.uuid[:8]}")
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_cli,
                    concepto=conc_c,
                    debe=None, haber=round(monto_c, 2),
                    cfdi=c,
                ))

        else:
            continue  # tipo desconocido

        polizas.append(pol)
        num += 1

    return polizas


# ── Pólizas de Diario ─────────────────────────────────────────────────────────

def generar_polizas_diario(
    cfdis:                  list["CFDI"],
    rfc_empresa:            str,
    num_inicio:             int,
    terceros:               list["Tercero"],
    reglas_prod:            list,           # list[ReglaProdServ]
    cta_cobrar:             str  = "",      # CxC default (emitidas)
    cta_pagar:              str  = "",      # CxP default (recibidas)
    cta_iva_tras:           str  = "",      # IVA Trasladado (pasivo)
    cta_iva_acre:           str  = "",      # IVA Acreditable (activo)
    cta_iva_ret:            str  = "",      # IVA Retenido (ambos lados)
    cta_isr_ret:            str  = "",      # ISR Retenido (ambos lados)
    cta_default_ingreso:    str  = "",      # Ingresos fallback
    cta_default_gasto:      str  = "",      # Gastos fallback
) -> list[Poliza]:
    """
    Genera Pólizas de Diario (Dr) por reconocimiento contable de CFDIs.

    CFDIs emitidas (empresa cobra) → Dr Clientes / Cr Ingresos + IVA Trasladado
    CFDIs recibidas (empresa paga) → Dr Gastos + IVA Acreditable / Cr Proveedores

    Solo procesa tipos I y E (facturas e notas de crédito). Excluye P y N.
    """
    from catalogo_prod_serv import cuenta_para_clave

    polizas:  list[Poliza] = []
    num = num_inicio
    rfc_u = rfc_empresa.strip().upper()

    _NO_DEF = _CUENTA_NO_DEFINIDA

    for cfdi in cfdis:
        if cfdi.tipo not in ("I", "E"):
            continue

        es_emitida = cfdi.rfc_emisor.upper() == rfc_u
        es_recibida = cfdi.rfc_receptor.upper() == rfc_u

        if not es_emitida and not es_recibida:
            continue

        subtotal_neto = round((cfdi.subtotal or 0) - (cfdi.descuento or 0), 2)
        iva16         = round(cfdi.iva16       or 0, 2)
        iva_ret       = round(cfdi.iva_retenido or 0, 2)
        isr_ret       = round(cfdi.isr_retenido or 0, 2)
        total         = round(cfdi.total        or 0, 2)
        dia           = _dia(cfdi.fecha_emision)

        cta_prod_ingreso = cuenta_para_clave(
            cfdi.clave_prod_serv, "ingreso", reglas_prod, cta_default_ingreso or _NO_DEF)
        cta_prod_gasto = cuenta_para_clave(
            cfdi.clave_prod_serv, "gasto",   reglas_prod, cta_default_gasto   or _NO_DEF)

        if es_emitida:
            # ── CFDI emitida: reconocemos ingreso ─────────────────────────────
            nombre_cli = cfdi.nombre_receptor or cfdi.rfc_receptor
            concepto   = _truncar(f"Factura — {nombre_cli}")
            pol = Poliza(tipo="Dr", numero=num, concepto=concepto, dia=dia)

            cta_cli = _cuenta_tercero(cfdi, False, terceros, cta_cobrar or _NO_DEF)

            # DEBE: Clientes (monto total que el cliente nos debe pagar)
            pol.partidas.append(PartidaPoliza(
                cuenta=cta_cli, concepto=_truncar(f"CxC {nombre_cli}"),
                debe=total, haber=None, cfdi=cfdi,
            ))
            # Si hay retenciones: también son activos por recuperar
            if iva_ret > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_iva_ret or _NO_DEF,
                    concepto="IVA Retenido x Recuperar",
                    debe=iva_ret, haber=None,
                ))
            if isr_ret > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_isr_ret or _NO_DEF,
                    concepto="ISR Retenido x Recuperar",
                    debe=isr_ret, haber=None,
                ))

            # HABER: Ingresos (subtotal neto)
            pol.partidas.append(PartidaPoliza(
                cuenta=cta_prod_ingreso,
                concepto=_truncar(f"Ingresos {cfdi.clave_prod_serv or 'S/C'}"),
                debe=None, haber=subtotal_neto,
            ))
            # HABER: IVA Trasladado (pasivo)
            if iva16 > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_iva_tras or _NO_DEF,
                    concepto="IVA Trasladado 16%",
                    debe=None, haber=iva16,
                ))

        else:
            # ── CFDI recibida: reconocemos gasto/compra ───────────────────────
            nombre_prov = cfdi.nombre_emisor or cfdi.rfc_emisor
            concepto    = _truncar(f"Gasto — {nombre_prov}")
            pol = Poliza(tipo="Dr", numero=num, concepto=concepto, dia=dia)

            cta_prov = _cuenta_tercero(cfdi, True, terceros, cta_pagar or _NO_DEF)

            # DEBE: Gastos (subtotal neto)
            pol.partidas.append(PartidaPoliza(
                cuenta=cta_prod_gasto,
                concepto=_truncar(f"Gasto {cfdi.clave_prod_serv or 'S/C'}"),
                debe=subtotal_neto, haber=None, cfdi=cfdi,
            ))
            # DEBE: IVA Acreditable
            if iva16 > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_iva_acre or _NO_DEF,
                    concepto="IVA Acreditable 16%",
                    debe=iva16, haber=None,
                ))

            # HABER: Proveedores (total)
            pol.partidas.append(PartidaPoliza(
                cuenta=cta_prov, concepto=_truncar(f"CxP {nombre_prov}"),
                debe=None, haber=total,
            ))
            # HABER: IVA Retenido a pagar (si aplica)
            if iva_ret > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_iva_ret or _NO_DEF,
                    concepto="IVA Retenido a pagar",
                    debe=None, haber=iva_ret,
                ))
            if isr_ret > 0:
                pol.partidas.append(PartidaPoliza(
                    cuenta=cta_isr_ret or _NO_DEF,
                    concepto="ISR Retenido a pagar",
                    debe=None, haber=isr_ret,
                ))

        polizas.append(pol)
        num += 1

    return polizas


# ── Serialización a filas COI ─────────────────────────────────────────────────

def polizas_a_filas(polizas: list[Poliza]) -> list[tuple]:
    """
    Convierte la lista de pólizas al formato de filas para Aspel COI.
    Cada tupla corresponde a una fila (10 columnas A-J).
    """
    filas: list[tuple] = []

    for pol in polizas:
        # Fila de encabezado
        filas.append((pol.tipo, pol.numero, pol.concepto[:120], pol.dia,
                      None, None, None, None, None, None))

        for partida in pol.partidas:
            # Fila de partida
            filas.append((
                None,
                partida.cuenta,
                partida.depto,
                partida.concepto[:120] if partida.concepto else None,
                partida.tipo_cambio,
                partida.debe,
                partida.haber,
                partida.centro_costo,
                partida.proyecto,
                None,
            ))

            # Filas de CFDI (si la partida tiene CFDI ligado)
            if partida.cfdi:
                c = partida.cfdi
                fecha_cfdi = _fecha_cfdi(c.fecha_timbrado or c.fecha_emision)
                filas.append((None, None, "INICIO_CFDI",
                               None, None, None, None, None, None, None))
                filas.append((
                    None, None,
                    fecha_cfdi,            # C: FechaCFDI (datetime para openpyxl)
                    c.serie or None,       # D: Serie
                    c.folio or None,       # E: Folio
                    c.rfc_emisor or None,  # F: RFC Emisor
                    c.rfc_receptor or None,# G: RFC Receptor
                    c.total or None,       # H: Monto CFDI
                    c.uuid or None,        # I: UUID
                    1,                     # J: Tipo (1=CFDI nacional)
                ))
                filas.append((None, None, "FIN_CFDI",
                               None, None, None, None, None, None, None))

        # Cierre de partidas
        filas.append((None, "FIN_PARTIDAS",
                      None, None, None, None, None, None, None, None))

    return filas


def _fecha_cfdi(fecha_str: str) -> Optional[datetime]:
    """Parsea la fecha del CFDI como objeto datetime para que Excel la trate como fecha."""
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(fecha_str[:len(fmt)], fmt[:len(fecha_str)])
        except (ValueError, TypeError):
            pass
    return None
