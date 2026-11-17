"""Herramienta 1 — Resumen de eventos (antes: main.py).

Calcula métricas generales y, al final, la tabla "IMPUTADOS por tipo de evento"
(la tercera tabla de la hoja de resumen del CMU).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from app_general.nucleo import Accion, ArchivoSpec, Entradas, Herramienta, Logger, Resultado

COL_TIPIFICACION = "TIPIFICACION"
COL_DETENIDOS = "DETENIDOS"
COL_CONTRAVENTORES = "CONTRAVENTORES"
COL_PC = "PC"
COL_TIPO_EVENTO = "TIPO EVENTO"
COL_CIERRE = "TIPO DE CIERRE"
REQUERIDAS = [COL_TIPIFICACION, COL_DETENIDOS, COL_CONTRAVENTORES, COL_PC]

# "colaboracion" con o sin tilde, ignorando mayúsculas y espacios.
PATRON_COLABORACION = r"^colaboraci[oó]n$"

# Cierres que cuentan como "imputados propios" (se compara en mayúsculas y sin espacios
# sobrantes). En el CMU de junio el valor es el primero; se deja también la variante
# con "PROPIO/S" por si algún mes viene escrita así.
CIERRES_IMPUTADOS = {
    "FINALIZA CON IMPUTADO/S",
    "FINALIZA CON IMPUTADO/S PROPIO/S",
    "FINALIZADA CON IMPUTADO/S PROPIO/S",
    "FINALIZADA CON IMPUTADO/S",
}

# Tipos de evento que se suman dentro de otro en la tabla de imputados.
EQUIVALENCIAS_EVENTO = {"ROBO AMA": "ROBO/HURTO"}


# ─────────────────────────────────────────────────────────────
#  Lectura robusta del libro
# ─────────────────────────────────────────────────────────────
def _leer_datos(ruta, hoja: str | int, log: Logger) -> pd.DataFrame:
    """Lee la hoja elegida; si no tiene las columnas esperadas, busca otra que sí.

    El CMU trae varias hojas (los datos están en 'Hoja1' pero la primera del libro
    puede ser una tabla dinámica), así que no conviene confiar en "la primera hoja".
    """
    with pd.ExcelFile(ruta) as libro:
        nombres = list(libro.sheet_names)
        elegida = nombres[hoja] if isinstance(hoja, int) else hoja
        for nombre in [elegida, *[n for n in nombres if n != elegida]]:
            encabezado = pd.read_excel(libro, sheet_name=nombre, nrows=0)
            columnas = {str(c).strip() for c in encabezado.columns}
            if all(c in columnas for c in REQUERIDAS):
                if nombre != elegida:
                    log.aviso(
                        f"⚠ La hoja '{elegida}' no tiene las columnas esperadas: "
                        f"se usa '{nombre}'."
                    )
                log.info(f"Hoja de datos: {nombre}")
                df = pd.read_excel(libro, sheet_name=nombre)
                df.columns = [str(c).strip() for c in df.columns]
                return df
    raise ValueError(
        f"Ninguna hoja tiene las columnas {REQUERIDAS}. Hojas revisadas: {nombres}"
    )


def _cortar_en_primera_fila_vacia(df: pd.DataFrame, log: Logger) -> pd.DataFrame:
    """Descarta lo que haya debajo de la primera fila totalmente vacía (notas al pie)."""
    vacias = df.isna().all(axis=1).to_numpy()
    if not vacias.any():
        return df
    pos = int(vacias.argmax())
    log.info(f"Se descartaron {len(df) - pos} filas bajo la primera fila vacía (notas al pie).")
    return df.iloc[:pos].reset_index(drop=True)


# ─────────────────────────────────────────────────────────────
#  Tabla: imputados por tipo de evento
# ─────────────────────────────────────────────────────────────
def tabla_imputados(df: pd.DataFrame, log: Logger) -> pd.DataFrame | None:
    """Réplica de la tabla 3 de la hoja de resumen.

    • Se consideran solo las filas con TIPO DE CIERRE 'FINALIZA CON IMPUTADO/S'.
    • CONTRAVENCION = suma de CONTRAVENTORES de esas filas.
    • Cada tipo de evento = suma de DETENIDOS (ROBO AMA se suma a ROBO/HURTO).
    • Se ordena de mayor a menor y se agrega 'Total general'.
    """
    faltan = [c for c in (COL_TIPO_EVENTO, COL_CIERRE) if c not in df.columns]
    if faltan:
        log.aviso(f"⚠ No se arma la tabla de imputados: faltan las columnas {faltan}.")
        return None

    cierre = df[COL_CIERRE].astype(str).str.strip().str.upper()
    filtradas = df[cierre.isin(CIERRES_IMPUTADOS)]
    if filtradas.empty:
        log.aviso(f"⚠ No hay filas con TIPO DE CIERRE en {sorted(CIERRES_IMPUTADOS)}.")
        return None
    log.info(f"Filas con cierre de imputados: {len(filtradas)}")

    tipo = (
        filtradas[COL_TIPO_EVENTO]
        .fillna("SIN TIPO DE EVENTO")
        .astype(str)
        .str.strip()
        .replace(EQUIVALENCIAS_EVENTO)
    )
    detenidos = filtradas.groupby(tipo)[COL_DETENIDOS].sum()
    contravencion = pd.Series({"CONTRAVENCION": filtradas[COL_CONTRAVENTORES].sum()})

    serie = pd.concat([contravencion, detenidos])
    serie = serie[serie > 0].sort_values(ascending=False, kind="stable")

    tabla = serie.astype(int).rename_axis("TIPO DE EVENTO").reset_index(name="IMPUTADOS")
    total = pd.DataFrame({"TIPO DE EVENTO": ["Total general"], "IMPUTADOS": [int(serie.sum())]})
    return pd.concat([tabla, total], ignore_index=True)


# ─────────────────────────────────────────────────────────────
#  Acción principal
# ─────────────────────────────────────────────────────────────
def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    ruta = entradas.ruta("eventos")
    log.titulo("RESUMEN DE EVENTOS")
    log.info(f"Leyendo {ruta.name}...")
    df = _leer_datos(ruta, entradas.hoja("eventos"), log)
    df = _cortar_en_primera_fila_vacia(df, log)
    log.info(f"Eventos leídos: {len(df)}")

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
    sin_presencia = int(((~es_colab) & ~tiene_pc).sum())

    log.ok(f"Detenidos totales (propios + colaboración): {detenidos_totales}")
    log.info(f"Intervenciones: {intervenciones}")
    log.info(f"Colaboraciones: {colaboraciones}")
    log.info(f"Detenidos en colaboración: {detenidos_colab}")
    log.info(f"Detenidos propios: {detenidos_propios}")
    log.info(f"Total de eventos: {total_eventos}")
    log.info(f"Eventos sin presencia policial ni colaboración: {sin_presencia}")

    tablas: dict[str, pd.DataFrame] = {}
    tabla = tabla_imputados(df, log)
    if tabla is not None:
        titulo = "Imputados por tipo de evento (cierre: FINALIZA CON IMPUTADO/S)"
        tablas[titulo] = tabla
        log.titulo(titulo.upper())
        for linea in tabla.to_string(index=False).splitlines():
            log.info(linea)

    return Resultado(
        mensaje="Resumen calculado.",
        metricas={
            "Total de eventos": total_eventos,
            "Intervenciones": intervenciones,
            "Colaboraciones": colaboraciones,
            "Detenidos totales": detenidos_totales,
            "Detenidos en colaboración": detenidos_colab,
            "Detenidos propios": detenidos_propios,
            "Sin presencia": sin_presencia,
        },
        tablas=tablas,
    )


# ─────────────────────────────────────────────────────────────
#  Exportación a Excel
# ─────────────────────────────────────────────────────────────
def _dar_formato(ws, df: pd.DataFrame) -> None:
    """Encabezado en negrita con color, columnas ajustadas y fila 'Total' resaltada."""
    relleno = PatternFill("solid", fgColor="3F51B5")
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = relleno
        celda.alignment = Alignment(horizontal="center")

    for i, columna in enumerate(df.columns, start=1):
        ancho = max([len(str(columna)), *(len(str(v)) for v in df.iloc[:, i - 1])])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = ancho + 3

    for fila in ws.iter_rows(min_row=2):
        if str(fila[0].value).lower().startswith("total"):
            for celda in fila:
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", fgColor="E8EAF6")
    ws.freeze_panes = "A2"


def exportar_excel(entradas: Entradas, log: Logger) -> Resultado:
    """Recalcula el resumen y guarda las métricas y la tabla en un Excel nuevo."""
    resultado = ejecutar(entradas, log)

    ruta = entradas.ruta("eventos")
    salida = ruta.parent / f"{ruta.stem}_RESUMEN.xlsx"

    hojas = {
        "Resumen": pd.DataFrame(
            {"INDICADOR": list(resultado.metricas), "VALOR": list(resultado.metricas.values())}
        )
    }
    for titulo, tabla in resultado.tablas.items():
        hojas["Imputados"] = tabla  # hay una sola tabla; si hubiera más, un nombre por tabla

    log.titulo("EXPORTANDO A EXCEL")
    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            for nombre, df in hojas.items():
                df.to_excel(writer, sheet_name=nombre, index=False)
                _dar_formato(writer.sheets[nombre], df)
    except PermissionError:
        raise RuntimeError(
            f"No se pudo escribir '{salida.name}': está abierto en Excel. "
            "Cerralo y volvé a intentar."
        ) from None

    log.ok(f"✓ Hojas guardadas: {', '.join(hojas)}")
    log.ok(f"✓ Archivo: {salida}")
    resultado.mensaje = "Resumen exportado a Excel."
    resultado.salidas = [salida]
    return resultado


HERRAMIENTA = Herramienta(
    id="resumen",
    nombre="Resumen de eventos",
    corto="Resumen",
    icono="SUMMARIZE",
    resumen="Cuenta eventos, colaboraciones, intervenciones y detenidos; arma la tabla de imputados.",
    pasos=(
        "Lee el Excel de eventos. Si la hoja elegida no tiene las columnas esperadas, "
        "busca sola la que sí las tiene (ej. 'Hoja1'), y descarta lo que haya debajo de "
        "la primera fila vacía (notas al pie).",
        "Convierte DETENIDOS y CONTRAVENTORES a números (lo vacío o texto cuenta 0).",
        "Marca como COLABORACION las filas cuya TIPIFICACION sea 'colaboración' "
        "(con o sin tilde).",
        "Intervención = fila que NO es colaboración y tiene contenido en la columna PC.",
        "Detenidos totales = CONTRAVENTORES + DETENIDOS; se separan en propios y en "
        "colaboración.",
        "Al final muestra la tabla IMPUTADOS por tipo de evento: filtra los cierres "
        "'FINALIZA CON IMPUTADO/S', agrupa los contraventores como CONTRAVENCION, suma "
        "los detenidos por TIPO EVENTO (ROBO AMA se suma a ROBO/HURTO) y ordena de mayor "
        "a menor con su total general.",
        "'Exportar a Excel' repite el cálculo y guarda las métricas y la tabla con formato "
        "en un archivo nuevo (hojas 'Resumen' e 'Imputados').",
    ),
    entradas=(
        ArchivoSpec(
            "eventos",
            "Excel de eventos (CMU)",
            "Libro con TIPIFICACION, DETENIDOS, PC, TIPO EVENTO…",
            pide_hoja=True,
        ),
    ),
    genera="'Calcular resumen' no genera archivos. 'Exportar a Excel' guarda "
    "<nombre del CMU>_RESUMEN.xlsx (hojas Resumen e Imputados) junto al archivo de origen.",
    acciones=(
        Accion("Calcular resumen", "PLAY_ARROW", ejecutar),
        Accion("Exportar a Excel", "FILE_DOWNLOAD", exportar_excel),
    ),
)
