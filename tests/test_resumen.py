import unittest
from datetime import time
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from app_general.herramientas import resumen
from app_general.nucleo import Entradas, Logger


class ResumenTests(unittest.TestCase):
    def test_guion_en_pc_se_considera_vacio(self):
        datos = pd.DataFrame(
            {
                "TIPIFICACION": ["Robo"] * 4,
                "DETENIDOS": [0] * 4,
                "CONTRAVENTORES": [0] * 4,
                "PC": ["-", " - ", "", "123"],
            }
        )
        entradas = Entradas(archivos={"eventos": Path("eventos.xlsx")})

        with patch.object(resumen, "_leer_datos", return_value=datos):
            resultado = resumen.ejecutar(entradas, Logger())

        self.assertEqual(resultado.metricas["Intervenciones"], 1)
        self.assertEqual(resultado.metricas["Sin presencia"], 3)

    def test_pc_numerico_se_considera_con_presencia(self):
        datos = pd.DataFrame(
            {
                "TIPIFICACION": ["Robo"] * 3,
                "DETENIDOS": [0] * 3,
                "CONTRAVENTORES": [0] * 3,
                "PC": [123, 456, 0],
            }
        )
        entradas = Entradas(archivos={"eventos": Path("eventos.xlsx")})

        with patch.object(resumen, "_leer_datos", return_value=datos):
            resultado = resumen.ejecutar(entradas, Logger())

        self.assertEqual(resultado.metricas["Intervenciones"], 3)
        self.assertEqual(resultado.metricas["Sin presencia"], 0)

    def test_pc_con_formato_hora_se_considera_con_presencia(self):
        datos = pd.DataFrame(
            {
                "TIPIFICACION": ["Robo"] * 4,
                "DETENIDOS": [0] * 4,
                "CONTRAVENTORES": [0] * 4,
                "PC": [time(8, 30), time(12, 45), None, "-"],
            }
        )
        entradas = Entradas(archivos={"eventos": Path("eventos.xlsx")})

        with patch.object(resumen, "_leer_datos", return_value=datos):
            resultado = resumen.ejecutar(entradas, Logger())

        self.assertEqual(resultado.metricas["Intervenciones"], 2)
        self.assertEqual(resultado.metricas["Sin presencia"], 2)
