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
    import complementos.presentismo as pres

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
    import complementos.presentismo as pres

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
    metricas: dict[str, str | int] = {"Personas": len(resumen)}
    for columna, etiqueta in (
        ("-LM-", "Lic. médicas"),
        ("-LLT-", "Llegadas tarde"),
        ("-AUS-", "Ausentes"),
    ):
        if columna in resumen.columns:
            metricas[etiqueta] = int(pd.to_numeric(resumen[columna], errors="coerce").sum())

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
    resumen="Resume licencias médicas, llegadas tarde y ausencias por persona y grupo.",
    pasos=(
        "Detecta cuántos días tiene el mes a partir del nombre del archivo "
        "(ej. 'MAYO' → 31 días, 'FEBRERO' → 28).",
        "Recorre todas las hojas del Excel buscando filas de personal 'contratado' y "
        "'nivel' (AUX.), y detecta el grupo (A-F, ADM) de cada bloque.",
        "Unifica los códigos de cada día (LXUC, LXM, FAC… → LM; SAN, A → AUS).",
        "Cuenta por persona: -LM-, -LLT-, -AUS- y -LA-.",
        "Genera el detalle completo y el resumen por persona.",
    ),
    entradas=(
        ArchivoSpec("presentismo", "Excel de presentismo", "Libro con una hoja por turno/grupo"),
    ),
    genera="DATOS/filtrado_presentismo_grupos.xlsx y DATOS/presentismo_resumen.xlsx.",
    acciones=(Accion("Procesar presentismo", "PLAY_ARROW", ejecutar),),
)
