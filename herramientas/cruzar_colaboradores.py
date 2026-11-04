"""Herramienta 2 — Cruzar colaboradores por SAE (antes: cruzar_colaboradores.py, PyQt6)."""

from __future__ import annotations

import pandas as pd

from app_general.nucleo import Accion, ArchivoSpec, Entradas, Herramienta, Logger, Resultado

COL_SAE = "SAE"
COLS_ARCHIVO2 = [f"COLABORADORES {i}" for i in range(1, 6)]  # nombres en el Archivo 2
COLS_SALIDA = [f"COLABORADOR {i}" for i in range(1, 6)]  # nombres que se insertan
POSICION_INSERCION = 7  # entre la columna 7 y la 8 del Archivo 1


def _a_texto(valor):
    """Convierte un valor de celda a texto limpio (conserva los vacíos)."""
    return str(valor).strip() if not pd.isna(valor) else valor


def _normalizar_sae(serie: pd.Series) -> pd.Series:
    """'48291057.0' → '48291057'. Excel a veces entrega los números como float."""

    def _limpiar(v):
        if pd.isna(v):
            return v
        v = str(v).strip()
        try:
            return str(int(float(v)))
        except ValueError:
            return v  # texto no numérico: se deja tal cual

    return serie.map(_limpiar)


def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    ruta1 = entradas.ruta("archivo1")
    ruta2 = entradas.ruta("archivo2")
    hoja2 = entradas.hoja("archivo2")
    salida = ruta1.parent / "resultado_colaboradores.xlsx"

    log.titulo("CRUZAR COLABORADORES POR SAE")
    log.info(f"Leyendo Archivo 1: {ruta1.name}")
    df1 = pd.read_excel(ruta1, header=0, converters={COL_SAE: _a_texto})
    df1.columns = [str(c).strip() for c in df1.columns]

    log.info(f"Leyendo Archivo 2: {ruta2.name} (hoja: {hoja2})")
    df2 = pd.read_excel(ruta2, sheet_name=hoja2, header=0, converters={COL_SAE: _a_texto})
    df2.columns = [str(c).strip() for c in df2.columns]

    log.info("Validando columnas...")
    faltan2 = [c for c in [COL_SAE, *COLS_ARCHIVO2] if c not in df2.columns]
    if faltan2:
        raise KeyError(
            f"Archivo 2 no tiene las columnas {faltan2}.\n"
            f"Disponibles: {', '.join(map(str, df2.columns))}"
        )
    if COL_SAE not in df1.columns:
        raise KeyError(
            f"Archivo 1 no tiene la columna '{COL_SAE}'.\n"
            f"Disponibles: {', '.join(map(str, df1.columns))}"
        )

    log.info("Normalizando SAE y construyendo tabla de cruce...")
    df1[COL_SAE] = _normalizar_sae(df1[COL_SAE])
    df2[COL_SAE] = _normalizar_sae(df2[COL_SAE])

    df2 = df2.dropna(subset=[COL_SAE])
    repetidos = int(df2[COL_SAE].duplicated().sum())
    if repetidos:
        log.aviso(f"⚠ {repetidos} SAE repetidos en Archivo 2: se usa la primera aparición.")
    lookup = df2[[COL_SAE, *COLS_ARCHIVO2]].drop_duplicates(subset=COL_SAE).set_index(COL_SAE)

    log.info("Mapeando colaboradores...")
    nuevas = pd.DataFrame(
        {nuevo: df1[COL_SAE].map(lookup[orig]) for orig, nuevo in zip(COLS_ARCHIVO2, COLS_SALIDA)},
        index=df1.index,
    )

    log.info(f"Insertando 5 columnas entre la posición {POSICION_INSERCION} y {POSICION_INSERCION + 1}...")
    df_final = pd.concat(
        [df1.iloc[:, :POSICION_INSERCION], nuevas, df1.iloc[:, POSICION_INSERCION:]], axis=1
    )

    df_final.to_excel(salida, index=False)

    total = len(df_final)
    con_colab = int(nuevas.notna().any(axis=1).sum())
    celdas = int(nuevas.notna().sum().sum())
    log.ok(f"✓ Archivo guardado: {salida}")
    log.info(f"Filas con al menos un colaborador: {con_colab} de {total}")

    return Resultado(
        mensaje="Cruce completado.",
        metricas={
            "Filas procesadas": total,
            "Con colaboradores": con_colab,
            "Sin coincidencia": total - con_colab,
            "Colaboradores cargados": celdas,
        },
        salidas=[salida],
    )


HERRAMIENTA = Herramienta(
    id="cruzar",
    nombre="Cruzar colaboradores",
    corto="Colabor.",
    icono="MERGE_TYPE",
    resumen="Agrega COLABORADOR 1 a 5 al listado de eventos buscando por número de SAE.",
    pasos=(
        "Lee el Archivo 1 (eventos diarios, encabezados en la fila 1) y el Archivo 2 "
        "(hoja que elijas, con COLABORADORES 1 a 5).",
        "Normaliza los SAE de ambos (quita decimales tipo '123.0' y espacios).",
        "Busca cada SAE del Archivo 1 en el Archivo 2 (si hay repetidos usa el primero).",
        "Inserta COLABORADOR 1..5 entre las columnas 7 y 8 del Archivo 1.",
        "Guarda 'resultado_colaboradores.xlsx' junto al Archivo 1.",
    ),
    entradas=(
        ArchivoSpec("archivo1", "Archivo 1 — eventos diarios", "Contiene la columna SAE"),
        ArchivoSpec(
            "archivo2",
            "Archivo 2 — colaboradores",
            "Contiene SAE y COLABORADORES 1..5",
            pide_hoja=True,
        ),
    ),
    genera="resultado_colaboradores.xlsx (en la carpeta del Archivo 1).",
    acciones=(Accion("Cruzar archivos", "PLAY_ARROW", ejecutar),),
)
