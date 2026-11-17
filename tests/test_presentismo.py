import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from app_general.complementos import presentismo as procesador
from app_general.herramientas import presentismo as herramienta
from app_general.nucleo import Entradas, Logger


def crear_detalle() -> pd.DataFrame:
    columnas = [0, 1, 2, "HOJA_ORIGEN", 6, 7, 8, "-LM-", "-LLT-", "-AUS-", "-LA-"]
    filas = [
        ["Contratado", "Persona LM", 101.0, "Turno A", "LM", None, None, 1, 0, 0, 0],
        ["Contratado", "Persona LM", "101", "Turno B", None, "LM", None, 1, 0, 0, 0],
        ["Contratado", "ANA PEREZ", None, "Turno A", "LLT", None, None, 0, 1, 0, 0],
        ["Contratado", "Ana Pérez", None, "Turno B", None, "P", None, 0, 0, 0, 0],
        ["Contratado", "Persona AUS", 303, "Turno A", "AUS", None, None, 0, 0, 1, 0],
        ["Contratado", "Persona AUS", 303, "Turno B", None, "AUS", None, 0, 0, 1, 0],
        ["Contratado", "Persona presente", 404, "Turno A", "P", "FA", None, 0, 0, 0, 0],
        ["Contratado", "Persona LA", 505, "Turno A", "LA", None, None, 0, 0, 0, 1],
        ["Contratado", "Persona mixta", 606, "Turno A", "LM", "P", None, 1, 0, 0, 0],
        ["Contratado", "Persona tarde", 707, "Turno A", "LLT", None, "LLT", 0, 2, 0, 0],
    ]
    return pd.DataFrame(filas, columns=columnas)


class PresentismoTests(unittest.TestCase):
    def test_resumen_consolida_personas_y_calcula_categorias(self):
        resumen = procesador._resumen_personas(crear_detalle(), dias_mes=3)

        self.assertEqual(len(resumen), 7)
        por_nombre = resumen.set_index("NOMBRE")
        self.assertEqual(por_nombre.loc["Persona LM", "-LM-"], 2)
        self.assertTrue(por_nombre.loc["Persona LM", "LM TODO EL MES"])
        self.assertTrue(por_nombre.loc["Persona AUS", "AUS TODO EL MES"])
        self.assertFalse(por_nombre.loc["ANA PEREZ / Ana Pérez", "LLT TODO EL MES"])
        self.assertTrue(por_nombre.loc["Persona tarde", "LLT TODO EL MES"])
        self.assertTrue(por_nombre.loc["Persona presente", "PRESENTE TODO EL MES"])
        self.assertEqual(por_nombre.loc["Persona LA", "-LA-"], 1)
        self.assertFalse(por_nombre.loc["Persona mixta", "LM TODO EL MES"])

    def test_resumen_vacio_conserva_columnas(self):
        resumen = procesador._resumen_personas(pd.DataFrame(), dias_mes=30)
        self.assertEqual(len(resumen), 0)
        self.assertIn("PRESENTE TODO EL MES", resumen.columns)

    def test_metricas_cuentan_personas_y_no_repeticiones(self):
        with tempfile.TemporaryDirectory() as directorio:
            carpeta = Path(directorio)
            resumen = procesador._resumen_personas(crear_detalle(), dias_mes=3)
            ruta_resumen = carpeta / "presentismo_resumen.xlsx"
            resumen.to_excel(ruta_resumen, index=False)
            ruta_filtrado = carpeta / "filtrado_presentismo_grupos.xlsx"
            ruta_filtrado.touch()

            def procesar(archivo_origen, carpeta_salida, log_fn, mostrar_dialogos):
                self.assertEqual(Path(archivo_origen).name, "presentismo_mayo.xlsx")
                self.assertEqual(Path(carpeta_salida), carpeta)
                self.assertIsNotNone(log_fn)
                self.assertFalse(mostrar_dialogos)
                return True, str(ruta_filtrado), str(ruta_resumen)

            entradas = Entradas(
                archivos={"presentismo": carpeta / "presentismo_mayo.xlsx"}
            )
            with (
                patch.object(herramienta, "DATOS_DIR", carpeta),
                patch.object(procesador, "procesar_presentismo", side_effect=procesar),
            ):
                resultado = herramienta.ejecutar(entradas, Logger())

        self.assertEqual(resultado.metricas["Personas únicas"], 7)
        self.assertEqual(resultado.metricas["Personas con lic. médica"], 2)
        self.assertEqual(resultado.metricas["Lic. prolongada (todo el mes)"], 1)
        self.assertEqual(resultado.metricas["Personas con llegada tarde"], 2)
        self.assertEqual(resultado.metricas["Llegada tarde todo el mes"], 1)
        self.assertEqual(resultado.metricas["Personas ausentes"], 1)
        self.assertEqual(resultado.metricas["Ausente todo el mes"], 1)
        self.assertEqual(resultado.metricas["Personas con licencia anual"], 1)
        self.assertEqual(resultado.metricas["Presentes todo el mes"], 1)


if __name__ == "__main__":
    unittest.main()
