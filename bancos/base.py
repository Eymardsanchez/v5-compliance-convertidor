"""
Clase base para todos los extractores de bancos.
Cada banco implementa su propio extractor heredando de BancoExtractor.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Optional, Callable


@dataclass
class Movimiento:
    fecha: str = ""
    folio: str = ""
    descripcion: str = ""
    deposito: Optional[float] = None
    retiro: Optional[float] = None
    saldo: Optional[float] = None

    def tipo(self) -> str:
        if self.deposito:
            return "DEPOSITO"
        if self.retiro:
            return "RETIRO"
        return "DESCONOCIDO"


@dataclass
class InfoCuenta:
    banco: str = ""
    numero_cuenta: str = ""
    clabe: str = ""
    titular: str = ""
    rfc: str = ""
    periodo: str = ""
    saldo_inicial: Optional[float] = None
    saldo_final: Optional[float] = None
    total_depositos: Optional[float] = None
    total_retiros: Optional[float] = None


class BancoExtractor(ABC):
    """Interfaz común para todos los extractores de bancos."""

    NOMBRE: str = ""
    COLOR_HEX: str = "333333"

    @classmethod
    @abstractmethod
    def detectar(cls, primera_pagina_texto: str) -> bool:
        """Retorna True si el texto de la primera página corresponde a este banco."""
        ...

    @abstractmethod
    def extraer(
        self,
        pdf_path: str,
        progress_callback: Optional[Callable[[int, int], None]] = None,
    ) -> tuple[list[Movimiento], InfoCuenta]:
        """
        Extrae movimientos e información de cuenta del PDF.
        Retorna (lista_movimientos, info_cuenta).
        """
        ...
