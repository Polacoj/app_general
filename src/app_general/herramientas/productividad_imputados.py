"""Herramienta 4 — Productividad imputados (antes: productividad_imputados.pyw).

Se reescribió con operaciones "vectorizadas" de pandas (máscaras booleanas) en lugar
de recorrer fila por fila con iterrows(): mismo resultado, mucho más rápido y corto.
"""

from __future__ import annotations

import unicodedata

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

CONT_COL = "CONTRAVENTORES"
DET_COL = "DETENIDOS"
TIPIFICACION_COL = "TIPIFICACION"
CIERRE_COL = "TIPO DE CIERRE"
BARRIO_COL = "BARRIO"
ORIGEN_COL = "ORIGEN"
GAP_COL = "GAP"
SAE_COL = "SAE"
ACL_COL = "ACLARACIONES"
GRUPO_COL = "GRUPO"
SISEP_COL = "SISEP/ANILLO"
OPERADOR_INDEX = 6  # el operador se toma por posición (columna 7)

CIERRE_PROPIO = "FINALIZA CON IMPUTADO/S PROPIO/S"
REQUERIDAS = [
    CONT_COL, DET_COL, TIPIFICACION_COL, CIERRE_COL,
    BARRIO_COL, ORIGEN_COL, GAP_COL, SAE_COL, ACL_COL,
]  # fmt: skip


def _quitar_acentos(valor) -> str:
    texto = str(valor)
    return "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))


def _extraer(
    df: pd.DataFrame, columna_id: str, ident: pd.Series, filtro: pd.Series
) -> pd.DataFrame:
    """Devuelve contravenciones + detenidos + intervenciones destacadas.

    columna_id: nombre de la 1.ª columna de salida ("OPERADOR" o "GRUPO").
    ident:      serie con el valor de esa columna para cada fila.
    filtro:     máscara extra (ej. solo filas SISEP); todo True para operadores.
    """
    propio = df[CIERRE_COL] == CIERRE_PROPIO
    cont = pd.to_numeric(df[CONT_COL], errors="coerce").fillna(0) >= 1
    det = pd.to_numeric(df[DET_COL], errors="coerce").fillna(0) >= 1
    tipif = df[TIPIFICACION_COL].fillna("").astype(str).str.strip().str.upper()
    aclar = df[ACL_COL].fillna("").map(lambda v: _quitar_acentos(v).strip().lower())
    destacado = aclar.str.contains("integ") | aclar.str.contains("dest")

    def armar(mascara: pd.Series, tipo) -> pd.DataFrame:
        m = mascara & filtro
        return pd.DataFrame(
            {
                columna_id: ident[m],
                "TIPO DE EVENTO": tipo[m] if isinstance(tipo, pd.Series) else tipo,
                "BARRIO": df.loc[m, BARRIO_COL],
                "ORIGEN": df.loc[m, ORIGEN_COL],
                "GAP": df.loc[m, GAP_COL],
                "SAE": df.loc[m, SAE_COL],
            }
        )

    tipo_detenido = tipif.map(lambda t: "ROBO/HURTO" if t == "ROBO/HURTO" else "DELITO")
    return pd.concat(
        [
            armar(cont & propio, "CONTRAVENCION"),
            armar(det & propio, tipo_detenido),
            armar(destacado, "INTERVENCION DESTACADA"),
        ],
        ignore_index=True,
    )


def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    ruta = entradas.ruta("eventos")
    log.titulo("PRODUCTIVIDAD DE IMPUTADOS")
    log.info(f"Leyendo {ruta.name}...")
    df = pd.read_excel(ruta, engine="openpyxl")
    log.ok(f"✓ {len(df)} filas leídas")

    faltan = [c for c in REQUERIDAS if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas: {faltan}")

    df[BARRIO_COL] = df[BARRIO_COL].astype("string").str.lower()

    # — Operadores —
    log.info("→ Eventos por operador...")
    if df.shape[1] > OPERADOR_INDEX:
        operador = df.iloc[:, OPERADOR_INDEX]
    else:
        operador = pd.Series(pd.NA, index=df.index)
    todos = pd.Series(True, index=df.index)
    res_operador = _extraer(df, "OPERADOR", operador, todos)
    log.ok(f"✓ {len(res_operador)} eventos de operador")

    # — SISEP —
    log.info("→ Eventos SISEP...")
    if SISEP_COL in df.columns and GRUPO_COL in df.columns:
        es_sisep = df[SISEP_COL].isin(["SISEP", "SISEP "])
        res_sisep = _extraer(df, "GRUPO", df[GRUPO_COL].fillna(""), es_sisep)
    else:
        log.aviso(f"⚠ Faltan las columnas '{SISEP_COL}' o '{GRUPO_COL}': se omite SISEP.")
        res_sisep = pd.DataFrame(
            columns=["GRUPO", "TIPO DE EVENTO", "BARRIO", "ORIGEN", "GAP", "SAE"]
        )
    log.ok(f"✓ {len(res_sisep)} eventos SISEP")

    salida_op = DATOS_DIR / "productividad_imputados.xlsx"
    salida_si = DATOS_DIR / "productividad_imputados_sisep.xlsx"
    res_operador.to_excel(salida_op, index=False, engine="openpyxl")
    res_sisep.to_excel(salida_si, index=False, engine="openpyxl")
    log.ok(f"✅ Guardados:\n  ‣ {salida_op}\n  ‣ {salida_si}")

    por_tipo = res_operador["TIPO DE EVENTO"].value_counts()
    metricas: dict[str, str | int] = {
        "Eventos operador": len(res_operador),
        "Eventos SISEP": len(res_sisep),
    }
    for tipo, cantidad in por_tipo.items():
        metricas[str(tipo).capitalize()] = int(cantidad)

    return Resultado("Procesamiento completado.", metricas, [salida_op, salida_si])


HERRAMIENTA = Herramienta(
    id="imputados",
    nombre="Productividad de imputados",
    corto="Imputados",
    icono="GAVEL",
    resumen="Lista contravenciones, delitos e intervenciones destacadas por operador y SISEP.",
    pasos=(
        "Lee el Excel de eventos y pasa BARRIO a minúsculas.",
        "CONTRAVENCION: CONTRAVENTORES ≥ 1 y TIPO DE CIERRE 'FINALIZA CON IMPUTADO/S PROPIO/S'.",
        "DETENIDOS ≥ 1 con el mismo cierre: es ROBO/HURTO si la TIPIFICACION lo dice, "
        "si no es DELITO.",
        "INTERVENCION DESTACADA: la columna ACLARACIONES menciona 'integrado' o 'destacado'.",
        "Repite todo solo para filas SISEP, identificando por GRUPO en vez de OPERADOR.",
        "Guarda un Excel para operadores y otro para SISEP.",
    ),
    entradas=(
        ArchivoSpec("eventos", "Excel de eventos (CMU)", "Con CONTRAVENTORES, DETENIDOS, TIPO DE CIERRE…"),
    ),
    genera="DATOS/productividad_imputados.xlsx y DATOS/productividad_imputados_sisep.xlsx.",
    acciones=(Accion("Procesar imputados", "PLAY_ARROW", ejecutar),),
)
