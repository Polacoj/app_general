"""Herramienta 6 — Presentismo (antes: presentismo_grafica.py)."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from app_general.nucleo import (
    DATOS_DIR,
    Accion,
    ArchivoSpec,
    Entradas,
    Herramienta,
    Logger,
    Resultado,
)


def ajustar_dias_presentismo(nombre_archivo: str) -> int:
    """Configura cuántos días tiene el mes según el NOMBRE del archivo.

    complementos/presentismo.py usa dos variables globales (`dias` y `d`) para saber
    cuántas columnas de días leer. Hay que fijarlas ANTES de procesar; en tu script
    original lo hacía `ejecutar_procesamiento()`, pero el flujo de
    productividad_general nunca lo hacía (siempre usaba 30 días).
    """
    import app_general.complementos.presentismo as pres

    nombre = nombre_archivo.lower()
    if any(mes in nombre for mes in pres.meses_31):
        dias, d = 31, 35
    elif "febrero" in nombre:
        dias, d = 28, 32
    else:
        dias, d = 30, 34
    pres.dias, pres.d = dias, d
    return dias


def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    import app_general.complementos.presentismo as pres

    ruta = entradas.ruta("presentismo")
    log.titulo("PRESENTISMO")
    dias = ajustar_dias_presentismo(ruta.name)
    log.info(f"Mes detectado por el nombre del archivo: {dias} días")

    ok, ruta_filtrado, ruta_resumen = pres.procesar_presentismo(
        str(ruta),
        carpeta_salida=DATOS_DIR,
        log_fn=log,
        mostrar_dialogos=False,  # sin ventanitas de Tk: el panel ya muestra el log
    )
    if not ok:
        raise RuntimeError(ruta_filtrado)

    resumen = pd.read_excel(ruta_resumen)
    metricas: dict[str, str | int] = {
        "Personas únicas": len(resumen),
        "Personas con lic. médica": int((resumen["-LM-"] > 0).sum()),
        "Lic. prolongada (todo el mes)": int(resumen["LM TODO EL MES"].sum()),
        "Personas con llegada tarde": int((resumen["-LLT-"] > 0).sum()),
        "Llegada tarde todo el mes": int(resumen["LLT TODO EL MES"].sum()),
        "Personas ausentes": int((resumen["-AUS-"] > 0).sum()),
        "Ausente todo el mes": int(resumen["AUS TODO EL MES"].sum()),
        "Personas con licencia anual": int((resumen["-LA-"] > 0).sum()),
        "Presentes todo el mes": int(resumen["PRESENTE TODO EL MES"].sum()),
    }

    return Resultado(
        mensaje="Presentismo procesado.",
        metricas=metricas,
        salidas=[Path(ruta_filtrado), Path(ruta_resumen)],
    )


HERRAMIENTA = Herramienta(
    id="presentismo",
    nombre="Presentismo",
    corto="Presentismo",
    icono="EVENT_AVAILABLE",
    resumen="Cuenta personas únicas con licencias, tardanzas, ausencias y presentismo mensual.",
    pasos=(
        "Detecta cuántos días tiene el mes a partir del nombre del archivo "
        "(ej. 'MAYO' → 31 días, 'FEBRERO' → 28).",
        "Recorre todas las hojas del Excel buscando filas de personal 'contratado' y "
        "'nivel' (AUX.), y detecta el grupo (A-F, ADM) de cada bloque.",
        "Unifica los códigos de cada día (LXUC, LXM, FAC… → LM; SAN, A → AUS).",
        "Consolida personas repetidas por LP/DNI (o nombre si falta el identificador).",
        "Cuenta personas con licencia médica, llegada tarde, ausencia y licencia anual.",
        "Identifica licencias médicas prolongadas, tardanzas o ausencias durante todos los días asignados, "
        "y personas presentes todos los días asignados (P o FA). Los días vacíos no cuentan.",
        "Genera el detalle completo y un resumen consolidado por persona.",
    ),
    entradas=(
        ArchivoSpec("presentismo", "Excel de presentismo", "Libro con una hoja por turno/grupo"),
    ),
    genera="DATOS/filtrado_presentismo_grupos.xlsx y DATOS/presentismo_resumen.xlsx.",
    acciones=(Accion("Procesar presentismo", "PLAY_ARROW", ejecutar),),
)
