"""Registro de herramientas.

Para agregar una nueva: crear herramientas/mi_herramienta.py con un objeto
HERRAMIENTA e incluirlo en la tupla de abajo. La interfaz hace el resto.
"""

from app_general.herramientas import (
    cruzar_colaboradores,
    diccionario_civiles,
    presentismo,
    productividad_general,
    productividad_imputados,
    resumen,
)
from app_general.nucleo import Herramienta

HERRAMIENTAS: tuple[Herramienta, ...] = (
    resumen.HERRAMIENTA,
    cruzar_colaboradores.HERRAMIENTA,
    productividad_general.HERRAMIENTA,
    productividad_imputados.HERRAMIENTA,
    diccionario_civiles.HERRAMIENTA,
    presentismo.HERRAMIENTA,
)
