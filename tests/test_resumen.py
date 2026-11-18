import unittest
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
