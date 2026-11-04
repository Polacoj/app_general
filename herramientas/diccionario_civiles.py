"""Herramienta 5 — Diccionario civiles (antes: diccionario_CIVILES.pyw, Flet).

Dos acciones:
  • Validar y reconciliar: reemplaza nombres por LP/DNI usando el Diccionario Maestro.
  • Revertir: vuelve de LP/DNI a nombres usando la hoja VALIDACION.

Todo lo que se genera va a una carpeta MODIFICADOS junto a cada archivo de origen:
los originales NO se tocan (ni siquiera el diccionario).
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from app_general.nucleo import (
    Accion,
    ArchivoSpec,
    Entradas,
    Herramienta,
    Logger,
    ProcesoIncompleto,
    Resultado,
)

COL_ID = "LP/DNI"
COL_NOMBRE = "APELLIDO Y NOMBRE"
COL_OPERADOR = "OPERADOR"
COLS_FISCALIZACION = [COL_OPERADOR]
COLS_CONTROL = [COL_OPERADOR, *[f"COLABORADORES {i}" for i in range(1, 6)]]
CARPETA_MODIFICADOS = "MODIFICADOS"
HOJA_SHEET1 = "SHEET1"
HOJA_VALIDACION = "VALIDACION"

# (clave de entrada, etiqueta para el log, columnas a reconciliar)
TRABAJOS = (
    ("fiscalizacion", "Fiscalización Diaria/Mensual", COLS_FISCALIZACION),
    ("control", "Control Diario Mes", COLS_CONTROL),
)


# ─────────────────────────────────────────────────────────────
#  Utilidades de texto / tablas
# ─────────────────────────────────────────────────────────────
@dataclass
class ResultadoColumna:
    evaluadas: int = 0
    coincidencias: int = 0
    no_encontrados: int = 0
    vacias: int = 0
    sin_id: int = 0
    nombres_nuevos: dict[str, list[str]] = field(default_factory=dict)
    nombres_sin_id: dict[str, list[str]] = field(default_factory=dict)

    @property
    def porcentaje(self) -> float:
        return 100.0 if self.evaluadas == 0 else self.coincidencias / self.evaluadas * 100.0


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df


def normalizar_nombre(valor) -> str:
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto == "" or texto.lower() == "nan":
        return ""
    return " ".join(texto.upper().split())


def _serie_id(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})


def construir_mapa_nombres(df_dic: pd.DataFrame) -> dict[str, str]:
    nombres = df_dic[COL_NOMBRE].map(normalizar_nombre)
    ids = _serie_id(df_dic[COL_ID])
    mask = nombres != ""
    return dict(zip(nombres[mask], ids[mask], strict=False))


def construir_mapa_inverso(df_val: pd.DataFrame) -> dict[str, str]:
    ids = _serie_id(df_val[COL_ID])
    nombres = df_val[COL_NOMBRE].fillna("").astype(str).str.strip()
    mask = ids != ""
    return dict(zip(ids[mask], nombres[mask], strict=False))


def _buscar_hoja(nombres: list[str], objetivo: str) -> str | None:
    objetivo = objetivo.strip().upper()
    return next((n for n in nombres if n.strip().upper() == objetivo), None)


def _leer(ruta: Path, hoja: str | int) -> pd.DataFrame:
    return normalizar_columnas(pd.read_excel(ruta, sheet_name=hoja, dtype=str))


def _carpeta_modificados(ruta: Path) -> Path:
    carpeta = ruta.parent / CARPETA_MODIFICADOS
    carpeta.mkdir(exist_ok=True)
    return carpeta


def ruta_procesado(ruta: Path) -> Path:
    return _carpeta_modificados(ruta) / f"{ruta.stem}_PROCESADO{ruta.suffix}"


def ruta_revertido(ruta_proc: Path) -> Path:
    """Recibe el *_PROCESADO (que ya vive en MODIFICADOS) y devuelve el *_REVERTIDO."""
    base = ruta_proc.stem.removesuffix("_PROCESADO")
    return ruta_proc.parent / f"{base}_REVERTIDO{ruta_proc.suffix}"


# ─────────────────────────────────────────────────────────────
#  Reconciliación nombre → ID
# ─────────────────────────────────────────────────────────────
def _ubicaciones(col: str, mask: pd.Series) -> list[str]:
    # +2: la fila 1 del Excel son los encabezados y pandas cuenta desde 0.
    return [f"{col}:fila {int(i) + 2}" for i in mask.index[mask]]


def aplicar_mapa(
    df: pd.DataFrame, columnas: list[str], mapa: dict[str, str]
) -> tuple[pd.DataFrame, ResultadoColumna]:
    df_mod = df.copy().reset_index(drop=True)
    stats = ResultadoColumna()

    for col in columnas:
        if col not in df_mod.columns:
            raise RuntimeError(f"La hoja no contiene la columna '{col}'.")

        norm = df_mod[col].map(normalizar_nombre)
        vacias = norm.eq("")
        ids = norm.map(mapa)
        en_mapa = ~vacias & ids.notna()
        sin_id = en_mapa & ids.eq("")
        coinciden = en_mapa & ids.ne("")
        no_encontrados = ~vacias & ids.isna()

        stats.vacias += int(vacias.sum())
        stats.coincidencias += int(coinciden.sum())
        stats.no_encontrados += int(no_encontrados.sum())
        stats.sin_id += int(sin_id.sum())
        stats.evaluadas += int((~vacias).sum())

        for nombre in norm[no_encontrados].unique():
            stats.nombres_nuevos.setdefault(str(nombre), []).extend(
                _ubicaciones(col, norm.eq(nombre) & no_encontrados)
            )
        for nombre in norm[sin_id].unique():
            stats.nombres_sin_id.setdefault(str(nombre), []).extend(
                _ubicaciones(col, norm.eq(nombre) & sin_id)
            )

        df_mod.loc[coinciden, col] = ids[coinciden].astype(str)

    return df_mod, stats


def _guardar_diccionario_actualizado(
    df_dic: pd.DataFrame, nombres_nuevos: list[str], ruta_dic: Path
) -> Path:
    """Copia el diccionario a MODIFICADOS y le agrega los nombres faltantes en Sheet1.

    Se copia el libro completo (así se conservan otras hojas como VALIDACION) y se
    reescribe solo Sheet1. El original queda intacto.
    """
    destino = _carpeta_modificados(ruta_dic) / f"{ruta_dic.stem}_ACTUALIZADO{ruta_dic.suffix}"
    shutil.copy2(ruta_dic, destino)

    nuevas = [{c: "" for c in df_dic.columns} | {COL_NOMBRE: n} for n in nombres_nuevos]
    df_out = pd.concat([df_dic, pd.DataFrame(nuevas)], ignore_index=True)
    df_out = df_out.where(df_out.notna(), "")

    wb = load_workbook(destino)
    nombre_hoja = _buscar_hoja(wb.sheetnames, HOJA_SHEET1)
    ws = wb[nombre_hoja] if nombre_hoja else wb.create_sheet("Sheet1")
    if ws.max_row:
        ws.delete_rows(1, ws.max_row)
    for i, fila in enumerate(dataframe_to_rows(df_out, index=False, header=True), start=1):
        for j, valor in enumerate(fila, start=1):
            ws.cell(row=i, column=j, value=valor)
    wb.save(destino)
    return destino


def _reconciliar(
    df: pd.DataFrame,
    df_dic: pd.DataFrame,
    columnas: list[str],
    ruta_origen: Path,
    ruta_dic: Path,
    etiqueta: str,
    log: Logger,
) -> tuple[Path, ResultadoColumna]:
    if COL_ID not in df_dic.columns or COL_NOMBRE not in df_dic.columns:
        raise RuntimeError(f"El Diccionario Maestro debe tener las columnas '{COL_ID}' y '{COL_NOMBRE}'.")

    log.info(f"Reconciliando {etiqueta}...")
    df_mod, stats = aplicar_mapa(df, columnas, construir_mapa_nombres(df_dic))
    log.info(
        f"{etiqueta}: {stats.evaluadas} celdas, {stats.coincidencias} coinciden, "
        f"{stats.no_encontrados} no encontradas, {stats.sin_id} sin ID → {stats.porcentaje:.1f}%"
    )

    if stats.porcentaje < 100.0:
        for nombre, ubic in stats.nombres_sin_id.items():
            log.aviso(f"⚠ Sin LP/DNI en el diccionario: '{nombre}' en {ubic}")
        partes = [f"Coincidencia incompleta ({stats.porcentaje:.1f}%)."]
        if stats.nombres_nuevos:
            destino = _guardar_diccionario_actualizado(df_dic, list(stats.nombres_nuevos), ruta_dic)
            for nombre, ubic in stats.nombres_nuevos.items():
                log.aviso(f"⚠ Nombre nuevo: '{nombre}' en {ubic}")
            log.ok(f"✔ Copia del diccionario con los faltantes: {destino}")
            partes.append("Se creó una copia del diccionario con los nombres faltantes.")
        partes.append("Completá la columna LP/DNI, usá esa copia como Diccionario Maestro y volvé a correr.")
        raise ProcesoIncompleto(" ".join(partes))

    salida = ruta_procesado(ruta_origen)
    df_mod.to_excel(salida, index=False)
    log.ok(f"✔ {etiqueta} guardada: {salida}")
    return salida, stats


# ─────────────────────────────────────────────────────────────
#  ACCIÓN 1: validar y reconciliar
# ─────────────────────────────────────────────────────────────
def validar_y_reconciliar(entradas: Entradas, log: Logger) -> Resultado:
    log.titulo("DICCIONARIO CIVILES — VALIDAR Y RECONCILIAR")
    ruta_dic = entradas.ruta("diccionario")
    trabajos = [(c, e, cols, entradas.opcional(c)) for c, e, cols in TRABAJOS if entradas.opcional(c)]
    if not trabajos:
        raise ValueError("Seleccioná al menos un archivo: Fiscalización o Control Diario.")

    log.info(f"Cargando diccionario: {ruta_dic.name}")
    df_dic = _leer(ruta_dic, entradas.hoja("diccionario"))

    cargados, errores = [], []
    for clave, etiqueta, columnas, ruta in trabajos:
        log.info(f"Cargando {etiqueta}: {ruta.name}")
        df = _leer(ruta, entradas.hoja(clave))
        faltan = [c for c in columnas if c not in df.columns]
        if faltan:
            errores.append(f"{etiqueta}: faltan columnas {faltan}. Hay: {sorted(df.columns)}")
        cargados.append((etiqueta, columnas, ruta, df))
    if errores:
        raise RuntimeError("\n".join(errores))
    log.ok("✓ Archivos cargados y columnas validadas")

    salidas: list[Path] = []
    metricas: dict[str, str | int] = {}
    for etiqueta, columnas, ruta, df in cargados:
        salida, stats = _reconciliar(df, df_dic, columnas, ruta, ruta_dic, etiqueta, log)
        salidas.append(salida)
        metricas[f"{etiqueta} — coincidencia"] = f"{stats.porcentaje:.0f}%"
        metricas[f"{etiqueta} — celdas"] = stats.evaluadas

    return Resultado("Reconciliación al 100 %.", metricas, salidas)


# ─────────────────────────────────────────────────────────────
#  ACCIÓN 2: revertir ID → nombre
# ─────────────────────────────────────────────────────────────
def _revertir(ruta_proc: Path, columnas: list[str], inverso: dict[str, str], etiqueta: str, log: Logger):
    log.info(f"Revirtiendo {etiqueta}: {ruta_proc.name}")
    with pd.ExcelFile(ruta_proc) as libro:
        hoja = libro.sheet_names[0]
    df = _leer(ruta_proc, hoja).reset_index(drop=True)

    reemplazados = vacias = 0
    sin_match: list[str] = []
    for col in columnas:
        if col not in df.columns:
            log.aviso(f"⚠ Columna '{col}' no encontrada, se omite.")
            continue
        valores = df[col].fillna("").astype(str).str.strip()
        vacio = valores.eq("") | valores.str.lower().eq("nan")
        match = ~vacio & valores.isin(inverso)
        sin = ~vacio & ~valores.isin(inverso)
        vacias += int(vacio.sum())
        reemplazados += int(match.sum())
        df.loc[match, col] = valores[match].map(inverso)
        sin_match += [f"{col}:fila {int(i) + 2} → '{v}'" for i, v in valores[sin].items()]

    log.info(f"  Reemplazados: {reemplazados} · Vacíos: {vacias}")
    for detalle in sin_match:
        log.aviso(f"⚠ Sin coincidencia en VALIDACION: {detalle}")

    salida = ruta_revertido(ruta_proc)
    df.to_excel(salida, index=False)
    log.ok(f"✔ Guardado: {salida}")
    return salida, reemplazados


def revertir(entradas: Entradas, log: Logger) -> Resultado:
    log.titulo("DICCIONARIO CIVILES — REVERTIR IDs → NOMBRES")
    ruta_dic = entradas.ruta("diccionario")

    with pd.ExcelFile(ruta_dic) as libro:
        hojas = list(libro.sheet_names)
    hoja_val = _buscar_hoja(hojas, HOJA_VALIDACION)
    if hoja_val is None:
        raise RuntimeError(f"No hay hoja '{HOJA_VALIDACION}' en el diccionario. Hojas: {hojas}")
    df_val = _leer(ruta_dic, hoja_val)
    for col in (COL_ID, COL_NOMBRE):
        if col not in df_val.columns:
            raise RuntimeError(f"La hoja '{hoja_val}' no tiene la columna '{col}'.")
    inverso = construir_mapa_inverso(df_val)
    log.info(f"Mapa inverso: {len(inverso)} entradas")

    salidas: list[Path] = []
    total = 0
    for clave, etiqueta, columnas in TRABAJOS:
        origen = entradas.opcional(clave)
        if not origen:
            continue
        procesado = ruta_procesado(origen)
        if not procesado.exists():
            raise RuntimeError(f"No existe {procesado.name}. Ejecutá primero 'Validar y reconciliar'.")
        salida, n = _revertir(procesado, columnas, inverso, etiqueta, log)
        salidas.append(salida)
        total += n
    if not salidas:
        raise ValueError("Seleccioná al menos un archivo: Fiscalización o Control Diario.")

    return Resultado("Reversión completada.", {"Celdas revertidas": total}, salidas)


HERRAMIENTA = Herramienta(
    id="civiles",
    nombre="Diccionario de civiles",
    corto="Civiles",
    icono="BADGE",
    resumen="Reemplaza nombres de OPERADOR/colaboradores por su LP/DNI, y revierte.",
    pasos=(
        "Carga el Diccionario Maestro (columnas LP/DNI y APELLIDO Y NOMBRE) y al menos un "
        "archivo: Fiscalización (columna OPERADOR) o Control Diario (OPERADOR + COLABORADORES 1-5).",
        "Normaliza cada nombre (mayúsculas, espacios) y lo busca en el diccionario.",
        "Si TODO coincide: guarda el archivo *_PROCESADO con los IDs en la carpeta MODIFICADOS.",
        "Si faltan nombres: crea una copia del diccionario con los nombres nuevos (en "
        "MODIFICADOS), para que completes su LP/DNI, y detiene el proceso.",
        "'Revertir': usa la hoja VALIDACION para volver de ID a nombre y guarda *_REVERTIDO.",
    ),
    entradas=(
        ArchivoSpec("diccionario", "Diccionario Maestro", "Con Sheet1 y hoja VALIDACION", pide_hoja=True),
        ArchivoSpec("fiscalizacion", "Fiscalización diaria/mensual", "Opcional", False, True),
        ArchivoSpec("control", "Control Diario Mes", "Opcional", False, True),
    ),
    genera="Carpeta MODIFICADOS (junto a cada archivo): *_PROCESADO, *_REVERTIDO y *_ACTUALIZADO.",
    acciones=(
        Accion("Validar y reconciliar", "PLAY_ARROW", validar_y_reconciliar),
        Accion("Revertir IDs → Nombres", "UNDO", revertir),
    ),
)
