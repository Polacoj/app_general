from . import (
    cruzar_colaboradores,
    diccionario_civiles,
    presentismo,
    productividad_general,
    productividad_imputados,
    resumen,
)
from ..nucleo import Herramienta

HERRAMIENTAS: tuple[Herramienta, ...] = (
    resumen.HERRAMIENTA,
    cruzar_colaboradores.HERRAMIENTA,
    productividad_general.HERRAMIENTA,
    productividad_imputados.HERRAMIENTA,
    diccionario_civiles.HERRAMIENTA,
    presentismo.HERRAMIENTA,
)
