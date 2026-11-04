"""Herramienta 1 — Resumen de eventos (antes: main.py)."""

from __future__ import annotations

import pandas as pd

from app_general.nucleo import Accion, ArchivoSpec, Entradas, Herramienta, Logger, Resultado

COL_TIPIFICACION = "TIPIFICACION"
COL_DETENIDOS = "DETENIDOS"
COL_CONTRAVENTORES = "CONTRAVENTORES"
COL_PC = "PC"
REQUERIDAS = [COL_TIPIFICACION, COL_DETENIDOS, COL_CONTRAVENTORES, COL_PC]

# "colaboracion" con o sin tilde, ignorando mayúsculas y espacios.
PATRON_COLABORACION = r"^colaboraci[oó]n$"


def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    ruta = entradas.ruta("eventos")
    log.titulo("RESUMEN DE EVENTOS")
    log.info(f"Leyendo {ruta.name}...")
    df = pd.read_excel(ruta, sheet_name=entradas.hoja("eventos"))
    log.info(f"Filas leídas: {len(df)}")

    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError(
            f"Faltan columnas: {faltan}\nColumnas detectadas: {list(df.columns)}"
        )

    # Números "sucios" (texto, vacíos) → 0, para poder sumar sin errores.
    for col in (COL_DETENIDOS, COL_CONTRAVENTORES):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    es_colab = (
        df[COL_TIPIFICACION]
        .astype(str)
        .str.strip()
        .str.lower()
        .str.match(PATRON_COLABORACION)
    )
    tiene_pc = df[COL_PC].notna() & (df[COL_PC].astype(str).str.strip() != "")

    total_eventos = len(df)
    detenidos_totales = int(df[COL_CONTRAVENTORES].sum() + df[COL_DETENIDOS].sum())
    colaboraciones = int(es_colab.sum())
    detenidos_colab = int(
        df.loc[es_colab, COL_CONTRAVENTORES].sum() + df.loc[es_colab, COL_DETENIDOS].sum()
    )
    detenidos_propios = detenidos_totales - detenidos_colab
    intervenciones = int(((~es_colab) & tiene_pc).sum())

    log.ok(f"Detenidos totales (propios + colaboración): {detenidos_totales}")
    log.info(f"Intervenciones: {intervenciones}")
    log.info(f"Colaboraciones: {colaboraciones}")
    log.info(f"Detenidos en colaboración: {detenidos_colab}")
    log.info(f"Detenidos propios: {detenidos_propios}")
    log.info(f"Total de eventos: {total_eventos}")

    return Resultado(
        mensaje="Resumen calculado.",
        metricas={
            "Total de eventos": total_eventos,
            "Intervenciones": intervenciones,
            "Colaboraciones": colaboraciones,
            "Detenidos totales": detenidos_totales,
            "Detenidos en colaboración": detenidos_colab,
            "Detenidos propios": detenidos_propios,
        },
    )


HERRAMIENTA = Herramienta(
    id="resumen",
    nombre="Resumen de eventos",
    corto="Resumen",
    icono="SUMMARIZE",
    resumen="Cuenta eventos, colaboraciones, intervenciones y detenidos de un Excel.",
    pasos=(
        "Lee el Excel de eventos y verifica que existan las columnas TIPIFICACION, "
        "DETENIDOS, CONTRAVENTORES y PC.",
        "Convierte DETENIDOS y CONTRAVENTORES a números (lo vacío o texto cuenta 0).",
        "Marca como COLABORACION las filas cuya TIPIFICACION sea 'colaboración' "
        "(con o sin tilde).",
        "Intervención = fila que NO es colaboración y tiene contenido en la columna PC.",
        "Detenidos totales = CONTRAVENTORES + DETENIDOS; se separan en propios y en "
        "colaboración.",
    ),
    entradas=(
        ArchivoSpec("eventos", "Excel de eventos", "Libro con TIPIFICACION, DETENIDOS, PC…"),
    ),
    genera="No genera archivos: solo muestra las métricas y el log.",
    acciones=(Accion("Calcular resumen", "PLAY_ARROW", ejecutar),),
)
