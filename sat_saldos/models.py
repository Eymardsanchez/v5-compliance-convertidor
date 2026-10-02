"""Modelo de datos (dataclasses). Referencia: blueprint sección 6."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class OrigenCompensacion:
    """Bloque 'Origen del Saldo a favor / Pago de lo indebido' de un concepto."""
    tipo_origen: str                    # "Saldo a favor" | "Pago de lo indebido"
    origen_ejercicio: Optional[int] = None
    origen_concepto: Optional[str] = None
    saldo_aplicar: Optional[float] = None
    op_origen: Optional[str] = None
    monto_original: Optional[float] = None
    remanente_hist: Optional[float] = None   # remanente histórico antes de aplicar
    remanente_act: Optional[float] = None    # remanente actualizado (INPC) antes de aplicar
    fecha_saldo: Optional[str] = None


@dataclass
class Concepto:
    nombre: str                         # texto tal cual del acuse
    a_cargo: Optional[float] = None
    a_favor: Optional[float] = None
    compensaciones: Optional[float] = None   # positivo = monto compensado
    cantidad_cargo: Optional[float] = None
    cantidad_pagar: Optional[float] = None
    origen: Optional[OrigenCompensacion] = None


@dataclass
class Declaracion:
    archivo: str
    rfc: str
    razon_social: str
    es_anual: bool
    ejercicio: Optional[int]
    periodo: Optional[int]              # 1-12 mensual, 0 = del ejercicio (anual)
    tipo: str                           # "Normal" | "Complementaria"
    subtipo: str                        # normal|complementaria|modificacion|no-presentada|dejar-sin-efecto
    fecha_presentacion: Optional[datetime]
    operacion: Optional[str]            # clave única
    version: Optional[str]
    conceptos: list[Concepto] = field(default_factory=list)


@dataclass
class OrigenSaldo:
    """Resultado del análisis ISR: un saldo a favor (origen) y sus aplicaciones."""
    etiqueta: str                       # "Anual 2023", "Pago indebido jun-2023", ...
    ejercicio: Optional[int]
    monto_original: float
    fecha_origen: Optional[str]
    op_origen: Optional[str]
    aplicaciones: list = field(default_factory=list)  # list[dict]
    total_aplicado: float = 0.0
    remanente_nominal: float = 0.0
    remanente_actualizado_ultimo: Optional[float] = None  # del último acuse (informativo)
    estatus: str = ""


@dataclass
class ResultadoIVA:
    ejercicio: int
    periodo: int
    a_favor: float
    a_cargo: float
    cantidad_pagar: float
    tipo: str
    operacion: str


@dataclass
class ConceptoPago:
    """Un impuesto dentro del acuse de pago (línea de captura)."""
    nombre: str
    a_cargo: float = 0.0            # impuesto a cargo del periodo (contribución)
    actualizacion: float = 0.0     # parte actualizada (Art. 17-A CFF)
    recargos: float = 0.0          # recargos por mora (Art. 21 CFF)
    cantidad_a_cargo: float = 0.0  # a_cargo + actualizacion + recargos
    cantidad_a_pagar: float = 0.0  # tras acreditamientos/compensaciones


@dataclass
class AcusePago:
    """Acuse 'INFORMACIÓN REGISTRADA DE PAGOS' con línea de captura y pago recibido.

    Permite verificar si la declaración realmente se pagó y desglosar la parte
    actualizada y los recargos del importe pagado.
    """
    archivo: str
    rfc: str
    razon_social: str
    ejercicio: Optional[int]
    periodo: Optional[int]                 # 1-12 mensual
    tipo: str                              # "Normal" | "Complementaria"
    subtipo: str                           # normal | complementaria | modificacion | ...
    operacion: Optional[str]
    fecha_presentacion: Optional[datetime]
    vencimiento: Optional[str]             # fecha límite de pago
    # Línea de captura
    linea_captura: Optional[str] = None
    importe_a_pagar: float = 0.0           # importe total a pagar (autoritativo)
    vigente_hasta: Optional[str] = None
    # Desglose (suma de conceptos)
    contribucion: float = 0.0
    actualizacion: float = 0.0
    recargos: float = 0.0
    conceptos: list = field(default_factory=list)   # list[ConceptoPago]
    # Pago recibido en la institución de crédito
    institucion: Optional[str] = None
    fecha_pago: Optional[str] = None
    importe_pagado: Optional[float] = None
    medio_pago: Optional[str] = None
    operacion_bancaria: Optional[str] = None

    @property
    def pagado(self) -> bool:
        """True si hay comprobante de pago recibido (fecha e importe)."""
        return bool(self.fecha_pago) and (self.importe_pagado or 0) > 0


@dataclass
class DetalleIVA:
    """Determinación de IVA extraída de la DECLARACIÓN DETALLADA (vista previa).

    A diferencia del acuse, esta trae el renglón 'Acreditamiento del saldo a favor
    de periodos anteriores', que permite reconstruir el remanente acumulado real.
    """
    archivo: str
    rfc: str
    ejercicio: Optional[int]
    periodo: Optional[int]
    operacion: Optional[str]
    fecha_presentacion: Optional[datetime]
    subtipo: str
    a_cargo_periodo: Optional[float]        # CANTIDAD A CARGO (IVA a cargo del periodo)
    a_favor_periodo: Optional[float]        # IMPUESTO A FAVOR (saldo a favor generado)
    acreditamiento_anterior: float          # ACREDITAMIENTO DEL SALDO A FAVOR DE PERIODOS ANTERIORES
    impuesto_a_cargo_final: Optional[float] # IMPUESTO A CARGO (tras acreditamiento)
    cantidad_a_pagar: Optional[float]
