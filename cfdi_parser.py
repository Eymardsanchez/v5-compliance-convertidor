"""
cfdi_parser.py — Lector de CFDIs mexicanos (XML SAT)
Soporta CFDI 3.3 y 4.0. Extrae todos los campos fiscales relevantes.
"""

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

_NS_CFDI3   = "http://www.sat.gob.mx/cfd/3"
_NS_CFDI4   = "http://www.sat.gob.mx/cfd/4"
_NS_TFD     = "http://www.sat.gob.mx/TimbreFiscalDigital"
_NS_PAGO20  = "http://www.sat.gob.mx/Pagos20"
_NS_PAGO10  = "http://www.sat.gob.mx/Pagos"    # Complemento de Pago 1.0 (CFDI 3.3)
_NS_NOM12   = "http://www.sat.gob.mx/nomina12"

_RFC_RE = re.compile(r"^[A-ZÑ&]{3,4}[0-9]{6}[A-Z0-9]{3}$", re.IGNORECASE)


def _f(val, default: float = 0.0) -> float:
    try:
        return round(float(val or 0), 6)
    except (TypeError, ValueError):
        return default


def _s(val, default: str = "") -> str:
    return str(val or default).strip()


@dataclass
class CFDI:
    uuid:            str   = ""
    version:         str   = ""
    tipo:            str   = ""   # I=Ingreso E=Egreso P=Pago N=Nómina T=Traslado
    fecha_emision:   str   = ""
    fecha_timbrado:  str   = ""
    serie:           str   = ""
    folio:           str   = ""
    rfc_emisor:      str   = ""
    nombre_emisor:   str   = ""
    rfc_receptor:    str   = ""
    nombre_receptor: str   = ""
    subtotal:        float = 0.0
    descuento:       float = 0.0
    iva16:           float = 0.0
    iva_retenido:    float = 0.0
    isr_retenido:    float = 0.0
    total:           float = 0.0
    moneda:          str   = "MXN"
    archivo:         str   = ""
    # Complemento de pago tipo P (pago20:Totales/@MontoTotalPagos)
    fecha_pago:      str   = ""
    # Complemento de nómina tipo N (nomina12:Nomina)
    # total = TotalPercepciones + TotalOtrosPagos - TotalDeducciones (neto empleado)
    nom_percepciones: float = 0.0
    nom_deducciones:  float = 0.0
    nom_otros_pagos:  float = 0.0
    # CFDIs relacionados: UUIDs a los que este CFDI reduce su importe.
    # Aplica a tipo "E" (nota de crédito) y "P" (pago que liquida facturas).
    cfdi_relacionados:     list = field(default_factory=list)  # UUIDs reducidos (TipoRelacion 01/03)
    cfdi_sustituye:        list = field(default_factory=list)  # UUIDs sustituidos (TipoRelacion 04)
    cfdi_facturas_pagadas: list = field(default_factory=list)  # UUIDs liquidados por tipo P
    # TODAS las relaciones tal cual vienen en el XML: [(TipoRelacion, UUID)].
    # Las listas de arriba filtran por tipo (01/03 y 04) porque la conciliación
    # solo actúa sobre esas; aquí no se filtra nada, así el reporte puede mostrar
    # también el 07 (aplicación de anticipo) y los demás del catálogo.
    cfdi_rel_todos:        list = field(default_factory=list)
    # Campos fiscales (CFDI 4.0 obligatorios; 3.3 parcial)
    forma_pago:       str   = ""   # FormaPago: "01" Efectivo,"03" TC,"99" Por definir…
    metodo_pago:      str   = ""   # MetodoPago: "PUE" pago único / "PPD" parcialidades
    regimen_emisor:   str   = ""   # Emisor > RegimenFiscal
    regimen_receptor: str   = ""   # Receptor > RegimenFiscalReceptor
    uso_cfdi:         str   = ""   # Receptor > UsoCFDI
    clave_prod_serv:  str   = ""   # ClaveProdServ del primer Concepto (para pólizas Diario)
    descripcion_concepto: str = ""  # Descripcion del primer Concepto
    # TODOS los conceptos, en el orden del XML (para el reporte de conciliación)
    conceptos_clave:  list  = field(default_factory=list)  # ClaveProdServ de cada Concepto
    conceptos_desc:   list  = field(default_factory=list)  # Descripcion de cada Concepto
    # IVA desglosado por tasa (Base e Importe de cada TasaOCuota)
    iva_base_16:     float = 0.0   # Base gravada al 16%
    iva_base_8:      float = 0.0   # Base gravada al 8% (fronteriza)
    iva_base_0:      float = 0.0   # Base gravada al 0%
    iva_importe_8:   float = 0.0   # IVA trasladado al 8%
    iva_exento_base: float = 0.0   # Base de actos exentos (TipoFactor=Exento)

    @property
    def _fecha_ref(self) -> str:
        """Fecha de referencia para conciliación: pago si existe, si no emisión."""
        return self.fecha_pago or self.fecha_emision

    def _parse_fecha_ref(self) -> Optional[datetime]:
        """Parsea _fecha_ref aceptando ISO (YYYY-MM-DD) y formato banco (DD/MM/YYYY).
        Cachea el resultado: la fecha del CFDI no cambia y `mes`/`anio` se consultan
        miles de veces en la conciliación (era un hot spot de strptime)."""
        cached = getattr(self, "_fref_cache", False)
        if cached is not False:
            return cached
        ref = self._fecha_ref[:10]
        result: Optional[datetime] = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                result = datetime.strptime(ref, fmt)
                break
            except ValueError:
                pass
        self._fref_cache = result
        return result

    @property
    def mes(self) -> int:
        m = getattr(self, "_mes_cache", -1)
        if m != -1:
            return m
        dt = self._parse_fecha_ref()
        m = dt.month if dt else 0
        self._mes_cache = m
        return m

    @property
    def anio(self) -> int:
        a = getattr(self, "_anio_cache", -1)
        if a != -1:
            return a
        dt = self._parse_fecha_ref()
        a = dt.year if dt else 0
        self._anio_cache = a
        return a

    def es_rfc_valido(self, rfc: str) -> bool:
        return bool(_RFC_RE.match(rfc.strip()))


def parsear_cfdi(archivo: str | Path) -> Optional[CFDI]:
    """Lee un archivo .xml y retorna un objeto CFDI o None si no es válido."""
    try:
        tree = ET.parse(str(archivo))
        root = tree.getroot()
    except Exception:
        return None

    tag = root.tag
    ns = _NS_CFDI4 if "cfd/4" in tag else _NS_CFDI3

    c = CFDI(archivo=str(archivo))
    c.version        = _s(root.get("Version", root.get("version")))
    c.tipo           = _s(root.get("TipoDeComprobante"))
    c.fecha_emision  = _s(root.get("Fecha", ""))[:19]
    c.serie          = _s(root.get("Serie"))
    c.folio          = _s(root.get("Folio"))
    c.moneda         = _s(root.get("Moneda", "MXN"))
    c.subtotal       = _f(root.get("SubTotal", root.get("subTotal")))
    c.descuento      = _f(root.get("Descuento", root.get("descuento")))
    c.total          = _f(root.get("Total",    root.get("total")))
    c.forma_pago     = _s(root.get("FormaPago"))
    c.metodo_pago    = _s(root.get("MetodoPago"))

    emisor = root.find(f"{{{ns}}}Emisor")
    if emisor is not None:
        c.rfc_emisor     = _s(emisor.get("Rfc", emisor.get("rfc"))).upper()
        c.nombre_emisor  = _s(emisor.get("Nombre", emisor.get("nombre")))
        c.regimen_emisor = _s(emisor.get("RegimenFiscal"))

    receptor = root.find(f"{{{ns}}}Receptor")
    if receptor is not None:
        c.rfc_receptor    = _s(receptor.get("Rfc", receptor.get("rfc"))).upper()
        c.nombre_receptor = _s(receptor.get("Nombre", receptor.get("nombre")))
        c.regimen_receptor= _s(receptor.get("RegimenFiscalReceptor"))
        c.uso_cfdi        = _s(receptor.get("UsoCFDI"))

    # Conceptos: el primero alimenta las pólizas Diario (no cambia); la lista
    # completa es la que el reporte de conciliación muestra separada por " | ".
    conceptos = root.find(f"{{{ns}}}Conceptos")
    if conceptos is not None:
        for i, con in enumerate(conceptos.findall(f"{{{ns}}}Concepto")):
            clave = _s(con.get("ClaveProdServ"))
            desc  = _s(con.get("Descripcion"))
            if i == 0:
                c.clave_prod_serv      = clave
                c.descripcion_concepto = desc
            c.conceptos_clave.append(clave)
            c.conceptos_desc.append(desc)

    impuestos = root.find(f"{{{ns}}}Impuestos")
    if impuestos is not None:
        traslados = impuestos.find(f"{{{ns}}}Traslados")
        if traslados is not None:
            for t in traslados.findall(f"{{{ns}}}Traslado"):
                if _s(t.get("Impuesto")) == "002":
                    tipo_factor = _s(t.get("TipoFactor")).lower()
                    base_t = _f(t.get("Base"))
                    imp_t  = _f(t.get("Importe"))
                    c.iva16 += imp_t
                    if tipo_factor == "exento":
                        c.iva_exento_base += base_t
                    else:
                        tasa = _f(t.get("TasaOCuota"))
                        if abs(tasa - 0.16) < 0.001:
                            c.iva_base_16 += base_t
                        elif abs(tasa - 0.08) < 0.001:
                            c.iva_base_8  += base_t
                            c.iva_importe_8 += imp_t
                        elif abs(tasa) < 0.001:
                            c.iva_base_0 += base_t
        retenciones = impuestos.find(f"{{{ns}}}Retenciones")
        if retenciones is not None:
            for r in retenciones.findall(f"{{{ns}}}Retencion"):
                imp = _s(r.get("Impuesto"))
                if imp == "002":
                    c.iva_retenido += _f(r.get("Importe"))
                elif imp == "001":
                    c.isr_retenido += _f(r.get("Importe"))

    # ── CfdiRelacionados ──────────────────────────────────────────────────────
    # TipoRelacion "01/03" = nota de crédito / devolución (reduce factura).
    # TipoRelacion "04"    = sustitución (deja sin efecto a los UUID listados).
    _TIPOS_REDUCCION = {"01", "03"}
    for rel_node in root.findall(f"{{{ns}}}CfdiRelacionados"):
        tipo_rel = _s(rel_node.get("TipoRelacion"))
        for rel in rel_node.findall(f"{{{ns}}}CfdiRelacionado"):
            uuid_rel = _s(rel.get("UUID")).upper()
            if not uuid_rel:
                continue
            if (tipo_rel, uuid_rel) not in c.cfdi_rel_todos:
                c.cfdi_rel_todos.append((tipo_rel, uuid_rel))
            if tipo_rel == "04":
                if uuid_rel not in c.cfdi_sustituye:
                    c.cfdi_sustituye.append(uuid_rel)
            elif tipo_rel in _TIPOS_REDUCCION:
                if uuid_rel not in c.cfdi_relacionados:
                    c.cfdi_relacionados.append(uuid_rel)

    complemento = root.find(f"{{{ns}}}Complemento")
    if complemento is not None:
        # ── Timbre fiscal ──────────────────────────────────────────────────
        tfd = complemento.find(f"{{{_NS_TFD}}}TimbreFiscalDigital")
        if tfd is not None:
            c.uuid           = _s(tfd.get("UUID")).upper()
            c.fecha_timbrado = _s(tfd.get("FechaTimbrado", ""))[:19]

        # ── Complemento de Pago (TipoDeComprobante = "P") ────────────────────
        # Versión 2.0 (pago20, CFDI 4.0) y 1.0 (pago10, CFDI 3.3).
        # El nodo raíz siempre tiene Total="0"; el monto real viene del complemento.
        if c.tipo == "P":
            for ns_p in (_NS_PAGO20, _NS_PAGO10):
                pagos_node = complemento.find(f"{{{ns_p}}}Pagos")
                if pagos_node is None:
                    continue

                # pago20 tiene <Totales MontoTotalPagos="..."/>
                totales = pagos_node.find(f"{{{ns_p}}}Totales")
                if totales is not None:
                    monto_total = _f(totales.get("MontoTotalPagos"))
                    if monto_total:
                        c.total = monto_total
                    c.iva16        = _f(totales.get("TotalTrasladosImpuestoIVA16"))
                    c.iva_retenido = _f(totales.get("TotalRetencionesImpuestoIVA"))
                    c.isr_retenido = _f(totales.get("TotalRetencionesImpuestoISR"))
                else:
                    # pago10: sumar Monto de cada elemento <Pago>
                    monto_sum = sum(
                        _f(p.get("Monto"))
                        for p in pagos_node.findall(f"{{{ns_p}}}Pago")
                    )
                    if monto_sum:
                        c.total = round(monto_sum, 2)

                # FechaPago: tomamos la más reciente entre todos los <Pago>
                fecha_pago_max = ""
                for pago in pagos_node.findall(f"{{{ns_p}}}Pago"):
                    fp = _s(pago.get("FechaPago", ""))[:10]
                    if fp > fecha_pago_max:
                        fecha_pago_max = fp
                if fecha_pago_max:
                    c.fecha_pago = fecha_pago_max

                # UUIDs de facturas liquidadas por este complemento (R003)
                for pago in pagos_node.findall(f"{{{ns_p}}}Pago"):
                    for docto in pago.findall(f"{{{ns_p}}}DoctoRelacionado"):
                        id_doc = _s(docto.get("IdDocumento")).upper()
                        if id_doc and id_doc not in c.cfdi_facturas_pagadas:
                            c.cfdi_facturas_pagadas.append(id_doc)
                break  # solo uno de los dos namespaces puede estar presente

        # ── Complemento de Nómina 1.2 (TipoDeComprobante = "N") ──────────────
        # El nodo raíz tiene Total="0"; el monto real es:
        #   TotalPercepciones + TotalOtrosPagos - TotalDeducciones
        # FechaPago del nodo Nomina se usa como fecha de referencia bancaria.
        if c.tipo == "N":
            nomina = complemento.find(f"{{{_NS_NOM12}}}Nomina")
            if nomina is not None:
                percepciones = _f(nomina.get("TotalPercepciones"))
                deducciones  = _f(nomina.get("TotalDeducciones"))
                otros_pagos  = _f(nomina.get("TotalOtrosPagos"))
                neto = round(percepciones + otros_pagos - deducciones, 2)
                if neto:
                    c.total            = neto
                    c.nom_percepciones = percepciones
                    c.nom_deducciones  = deducciones
                    c.nom_otros_pagos  = otros_pagos
                # FechaPago = cuándo se deposita en la cuenta del empleado.
                # Se guarda en DD/MM/YYYY para coincidir con el formato del
                # estado de cuenta bancario y facilitar la conciliación visual.
                fp_iso = _s(nomina.get("FechaPago", ""))[:10]
                if fp_iso:
                    try:
                        c.fecha_pago = datetime.strptime(fp_iso, "%Y-%m-%d").strftime("%d/%m/%Y")
                    except ValueError:
                        c.fecha_pago = fp_iso

    return c if c.uuid else None


def leer_cfdis_carpetas(*rutas: str | Path,
                        recursivo: bool = True) -> list[CFDI]:
    """Lee CFDIs de carpetas (recursivo) y/o archivos .xml individuales."""
    cfdis: list[CFDI] = []
    for ruta in rutas:
        p = Path(ruta)
        if p.is_file():
            if p.suffix.lower() == ".xml":
                cfdi = parsear_cfdi(p)
                if cfdi:
                    cfdis.append(cfdi)
        elif p.is_dir():
            patron = "**/*.xml" if recursivo else "*.xml"
            for xml_path in p.glob(patron):
                cfdi = parsear_cfdi(xml_path)
                if cfdi:
                    cfdis.append(cfdi)
    return cfdis


def clasificar_cfdis(cfdis: list[CFDI],
                     rfc_empresa: str
                     ) -> tuple[list[CFDI], list[CFDI], list[CFDI]]:
    """
    Clasifica CFDIs en Recibidos, Emitidos y Descartados según el RFC de la empresa.

    Regla de pertenencia (descarte si no cumple):
      Recibidos  → rfc_receptor == empresa  (empresa pagó a proveedor → RETIRO)
                   rfc_emisor   == empresa  y tipo="N" (nómina → RETIRO)
      Emitidos   → rfc_emisor   == empresa  y tipo != "N" (empresa cobró → DEPÓSITO)
      Descartados→ el RFC de la empresa NO aparece ni como emisor ni como receptor.
                   Estos XMLs pertenecen a otra empresa y se ignoran.

    Returns (recibidos, emitidos, descartados).
    El tercer valor permite al llamador mostrar cuántos XMLs fueron ignorados.
    """
    rfc = rfc_empresa.strip().upper()

    recibidos:   list[CFDI] = []
    emitidos:    list[CFDI] = []
    descartados: list[CFDI] = []

    for c in cfdis:
        es_emisor   = c.rfc_emisor   == rfc
        es_receptor = c.rfc_receptor == rfc
        es_nomina   = c.tipo == "N"

        if es_receptor:
            # rfc_receptor == empresa → empresa pagó → RETIRO
            recibidos.append(c)
        elif es_emisor and es_nomina:
            # rfc_emisor == empresa y nómina → empresa pagó sueldos → RETIRO
            recibidos.append(c)
        elif es_emisor:
            # rfc_emisor == empresa y no nómina → empresa cobró → DEPÓSITO
            emitidos.append(c)
        else:
            # El RFC de la empresa no aparece en este XML → no pertenece a esta empresa
            descartados.append(c)

    return recibidos, emitidos, descartados
