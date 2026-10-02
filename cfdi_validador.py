"""
cfdi_validador.py — CFDI Fiscal Validation Engine v1.0

Statuses de salida:
  VERDE    score ≥ 90  → Deducible y conciliable
  AMARILLO score ≥ 70  → Requiere revisión manual
  ROJO     score ≥  0  → No deducible fiscalmente
  NEGRO    score <   0 → Sin efecto fiscal (crítico)

Uso básico:
    from cfdi_validador import validar_cfdi, validar_lote
    r = validar_cfdi(cfdi)           # validación individual
    mapa = validar_lote(cfdis)       # dict uuid_lower → resultado (detecta batch)
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from cfdi_parser import CFDI

# ── Constantes ────────────────────────────────────────────────────────────────

_TASA_IVA = 0.16

# Regímenes fiscales SAT vigentes (personas morales y físicas)
_REGIMENES_VALIDOS = {
    "601", "603", "605", "606", "607", "608", "610", "611", "612",
    "614", "615", "616", "620", "621", "622", "623", "624", "625",
    "626", "628", "629", "630",
}

_USO_CFDI_GENERICOS = {"S01", ""}


# ── Resultado ─────────────────────────────────────────────────────────────────

@dataclass
class CFDIValidationResult:
    uuid:             str
    errores:          list[str] = field(default_factory=list)
    advertencias:     list[str] = field(default_factory=list)
    score:            int  = 100
    estatus:          str  = "VERDE"   # VERDE / AMARILLO / ROJO / NEGRO
    reglas_aplicadas: list[str] = field(default_factory=list)


# ── Helpers de validación ─────────────────────────────────────────────────────

def _iva_es_valido(cfdi: "CFDI") -> bool:
    """
    True si el IVA 16% declarado es coherente con el subtotal (±1 % o ≤ $0.10).
    Retorna True si el CFDI no aplica o no tiene IVA (exento / tasa 0).
    """
    if cfdi.tipo not in ("I", "E"):
        return True
    if cfdi.iva16 <= 0:
        return True
    base = cfdi.subtotal - cfdi.descuento
    if base <= 0:
        return False
    iva_esperado = round(base * _TASA_IVA, 2)
    diferencia   = abs(cfdi.iva16 - iva_esperado)
    return diferencia <= 0.10 or (iva_esperado and diferencia / iva_esperado <= 0.01)


def _regimen_es_valido_sat(codigo: str) -> bool:
    """True si el código de régimen existe en el catálogo SAT."""
    return not codigo or codigo in _REGIMENES_VALIDOS


def _regimen_coincide_empresa(cfdi: "CFDI", regimenes_empresa: set[str],
                               es_emitido: bool) -> tuple[bool, str]:
    """
    Verifica que el régimen del lado empresa coincida con el de la CSF.
    Para emitidos: se verifica regimen_emisor (empresa es el emisor).
    Para recibidos: se verifica regimen_receptor (empresa es el receptor).

    Retorna (ok, mensaje_informativo). Un régimen distinto NO bloquea la
    conciliación ni la deducibilidad — el CFDI ya pertenece al RFC de la
    empresa (por eso se validó su lado); el mensaje es solo una anotación
    informativa para el contador (ver R006 en validar_cfdi).
    """
    reg_cfdi = (cfdi.regimen_emisor if es_emitido else cfdi.regimen_receptor).strip()
    lado     = "emisor" if es_emitido else "receptor"

    if not reg_cfdi:
        return False, (
            f"Régimen fiscal del {lado} no especificado en el CFDI "
            "(informativo — el CFDI pertenece al RFC de la empresa)"
        )
    if reg_cfdi not in regimenes_empresa:
        return False, (
            f"Régimen del {lado} en el CFDI ({reg_cfdi}) difiere del régimen "
            f"registrado de la empresa ({', '.join(sorted(regimenes_empresa))}) "
            "— informativo; el CFDI pertenece al RFC de la empresa y sí concilia"
        )
    return True, ""


# ── Validador principal ───────────────────────────────────────────────────────

def validar_cfdi(
    cfdi: "CFDI",
    *,
    estatus_sat:            str       = "",    # "VIGENTE" | "CANCELADO"
    tiene_complemento_pago: bool      = False, # True si existe CFDI tipo P que liquida este UUID
    es_sustituido:          bool      = False, # True si otro CFDI TipoRelacion 04 apunta a este
    uuid_duplicado:         bool      = False, # True si el UUID aparece más de una vez
    regimenes_empresa:      set[str] | None = None, # claves SAT de la empresa (de su CSF)
    es_emitido:             bool      = False, # True=empresa es emisor; False=empresa es receptor
) -> CFDIValidationResult:
    """Valida un CFDI ya parseado. Los parámetros opcionales requieren contexto externo."""

    errores:      list[str] = []
    advertencias: list[str] = []
    reglas:       list[str] = []
    delta = 0

    # ── R001: CFDI no timbrado o cancelado en SAT ─────────────────────────────
    if not cfdi.uuid:
        errores.append("Sin UUID — CFDI no timbrado o ilegible")
        delta -= 100
        reglas.append("R001a")
    if estatus_sat.upper() == "CANCELADO":
        errores.append("CFDI cancelado en el SAT")
        delta -= 100
        reglas.append("R001b")

    # ── R002: RFC emisor / receptor inválido o faltante ───────────────────────
    if not cfdi.rfc_emisor:
        errores.append("RFC emisor faltante")
        delta -= 40
        reglas.append("R002a")
    if not cfdi.rfc_receptor:
        errores.append("RFC receptor faltante")
        delta -= 40
        reglas.append("R002b")

    # ── R003: PPD sin complemento de pago ────────────────────────────────────
    if (cfdi.tipo == "I"
            and cfdi.metodo_pago == "PPD"
            and not tiene_complemento_pago):
        errores.append(
            "MetodoPago PPD sin complemento de pago (tipo P) — "
            "CFDI no deducible hasta recibir el complemento"
        )
        delta -= 90
        reglas.append("R003")

    # ── R004: UsoCFDI genérico o sin clasificar ───────────────────────────────
    if cfdi.tipo in ("I", "E") and cfdi.uso_cfdi in _USO_CFDI_GENERICOS:
        advertencias.append(
            f"UsoCFDI genérico o vacío ({cfdi.uso_cfdi or 'sin especificar'}) — "
            "verificar que corresponde al uso real de la operación"
        )
        delta -= 10
        reglas.append("R004")

    # ── R005: IVA inconsistente (no acreditable) ─────────────────────────────
    if cfdi.tipo in ("I", "E") and not _iva_es_valido(cfdi):
        base = cfdi.subtotal - cfdi.descuento
        errores.append(
            f"IVA inconsistente — declarado: ${cfdi.iva16:,.2f}, "
            f"esperado: ${round(base * _TASA_IVA, 2):,.2f} — IVA no acreditable"
        )
        delta -= 60
        reglas.append("R005")

    # ── R006: Régimen fiscal incompatible con la empresa (CSF) ────────────────
    # Aplica a todos los tipos timbrados: I/E (facturas), P (complemento de
    # pago) y N (nómina). Para emitidos se verifica Emisor>RegimenFiscal;
    # para recibidos se verifica Receptor>RegimenFiscalReceptor.
    if cfdi.tipo in ("I", "E", "P", "N"):
        if regimenes_empresa:
            # Régimen del lado empresa vs su CSF. El CFDI ya pertenece al RFC
            # de la empresa (fue clasificado como emitido/recibido), así que un
            # régimen distinto NO lo vuelve "no deducible" ni bloquea la
            # conciliación: se registra como ANOTACIÓN INFORMATIVA (advertencia),
            # sin penalizar el score. (Cambio pedido: conciliar si pertenece al
            # RFC de la empresa y solo anotar el régimen como dato informativo.)
            ok, msg = _regimen_coincide_empresa(cfdi, regimenes_empresa, es_emitido)
            if not ok:
                advertencias.append(msg)
                reglas.append("R006")
        else:
            # Sin CSF cargada: solo verificar que los códigos existan en el catálogo SAT
            reg_e = cfdi.regimen_emisor.strip()
            reg_r = cfdi.regimen_receptor.strip()
            if (reg_e and not _regimen_es_valido_sat(reg_e)) or \
               (reg_r and not _regimen_es_valido_sat(reg_r)):
                errores.append(
                    f"Régimen fiscal no reconocido por SAT — "
                    f"emisor: '{reg_e}', receptor: '{reg_r}'"
                )
                delta -= 70
                reglas.append("R006")

    # ── R007a: CFDI sustituido (sin efecto fiscal) ────────────────────────────
    if es_sustituido:
        errores.append(
            "CFDI SIN EFECTO FISCAL — fue sustituido por otro CFDI "
            "(TipoRelacion 04): no puede deducirse"
        )
        delta -= 120
        reglas.append("R007a")

    # ── R007b: Este CFDI sustituye a otros (advertencia al contador) ──────────
    if cfdi.cfdi_sustituye:
        advertencias.append(
            f"Este CFDI sustituye a {len(cfdi.cfdi_sustituye)} CFDI(s) previos — "
            "verificar que los CFDI sustituidos no estén también contabilizados"
        )
        reglas.append("R007b")

    # ── R008: UUID duplicado en cartera ───────────────────────────────────────
    if uuid_duplicado:
        errores.append(
            "UUID duplicado — el mismo CFDI aparece más de una vez en el sistema: "
            "riesgo de doble deducción"
        )
        delta -= 80
        reglas.append("R008")

    # ── R009: Datos fiscales obligatorios faltantes ───────────────────────────
    if cfdi.tipo in ("I", "E"):
        faltantes = []
        if not cfdi.forma_pago:
            faltantes.append("FormaPago")
        if not cfdi.metodo_pago:
            faltantes.append("MetodoPago")
        if faltantes:
            advertencias.append(
                f"Campos fiscales faltantes: {', '.join(faltantes)} — "
                "CFDI incompleto para conciliación fiscal"
            )
            delta -= 20
            reglas.append("R009")

    # ── Score y estatus final ─────────────────────────────────────────────────
    score = max(-200, 100 + delta)

    if score >= 90:
        estatus = "VERDE"
    elif score >= 70:
        estatus = "AMARILLO"
    elif score >= 0:
        estatus = "ROJO"
    else:
        estatus = "NEGRO"

    return CFDIValidationResult(
        uuid=cfdi.uuid,
        errores=errores,
        advertencias=advertencias,
        score=score,
        estatus=estatus,
        reglas_aplicadas=reglas,
    )


# ── Validación por lote (detecta contexto batch) ──────────────────────────────

def validar_lote(
    cfdis: list["CFDI"],
    rfc_empresa: str = "",
) -> dict[str, CFDIValidationResult]:
    """
    Valida toda la cartera de CFDIs.

    rfc_empresa: RFC de la empresa que gestiona la conciliación.
      Si está registrada con su CSF, se usan sus regímenes para validar R006.
      Si es emitido (empresa == rfc_emisor) se valida regimen_emisor;
      si es recibido (empresa == rfc_receptor) se valida regimen_receptor.

    Contexto batch detectado automáticamente:
      R003 — PPD con complemento de pago en el lote
      R007 — CFDIs sustituidos por TipoRelacion 04
      R008 — UUIDs duplicados en la cartera
    """
    # Regímenes de la empresa desde su CSF (si está registrada)
    regimenes_empresa: set[str] | None = None
    rfc_upper = rfc_empresa.strip().upper()
    if rfc_upper:
        from empresas import obtener_regimenes
        regs = obtener_regimenes(rfc_upper)
        if regs:
            regimenes_empresa = set(regs)

    # R008: detectar duplicados
    uuid_counts = Counter(c.uuid.lower() for c in cfdis if c.uuid)
    duplicados  = {u for u, n in uuid_counts.items() if n > 1}

    # R003: UUIDs de facturas liquidadas por algún complemento de pago (tipo P)
    facturas_pagadas: set[str] = set()
    for c in cfdis:
        if c.tipo == "P":
            for uid in c.cfdi_facturas_pagadas:
                facturas_pagadas.add(uid.lower())

    # R007: UUIDs sustituidos por algún CFDI con TipoRelacion 04
    sustituidos: set[str] = set()
    for c in cfdis:
        for uid in c.cfdi_sustituye:
            sustituidos.add(uid.lower())

    return {
        c.uuid.lower(): validar_cfdi(
            c,
            tiene_complemento_pago=(c.uuid.lower() in facturas_pagadas),
            es_sustituido=(c.uuid.lower() in sustituidos),
            uuid_duplicado=(c.uuid.lower() in duplicados),
            regimenes_empresa=regimenes_empresa,
            # Emitido = la empresa es el emisor del CFDI
            es_emitido=(bool(rfc_upper) and c.rfc_emisor == rfc_upper),
        )
        for c in cfdis if c.uuid
    }
