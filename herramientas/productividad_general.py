"""Herramienta 3 — Productividad general (antes: productividad_general.pyw, Tkinter).

Reutiliza tus clases de `complementos/` tal cual. Lo único nuevo es cómo se las
"enchufa" al panel: se les pasa el Logger como `log_fn`.
"""

from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import pandas as pd

from app_general.herramientas.presentismo import ajustar_dias_presentismo
from app_general.nucleo import (
    DATOS_DIR,
    Accion,
    ArchivoSpec,
    Entradas,
    Herramienta,
    Logger,
    Resultado,
)

# Posiciones (base 1) de las columnas a extraer del CMU. Los "" son columnas vacías
# que se agregan para respetar el formato de la plantilla.
COLUMNAS_A_EXTRAER = [1, 2, 7, 16, "", "", 20, 32, 13, 14]


# ─────────────────────────────────────────────────────────────
#  PASO A: archivo unificado a partir del CMU
# ─────────────────────────────────────────────────────────────
def _generar_unificado(ruta: Path, log: Logger) -> tuple[Path, dict[str, int]]:
    from complementos.colaboradores import ColaboradoresProcessor
    from complementos.contraventores_detenidos import ContraventoresProcessor
    from complementos.integrados import IntegradosProcessor

    log.titulo("📊 PASO 1: CARGA Y LIMPIEZA")
    df = pd.read_excel(ruta, engine="openpyxl")
    log.ok(f"✓ Archivo cargado: {df.shape[0]} filas × {df.shape[1]} columnas")

    # Cortar en la primera fila totalmente vacía.
    df_tmp = df.replace(r"^\s*$", pd.NA, regex=True)
    vacias = df_tmp.isna().all(axis=1)
    if vacias.any():
        pos = df_tmp.index.get_loc(vacias.idxmax())
        df = df.iloc[:pos].reset_index(drop=True)
        log.ok(f"✓ Filas truncadas en la primera fila vacía: quedan {pos}")
    else:
        log.info("No hay filas vacías para truncar")

    # Limpiar GAP y SAE: quedarse con lo anterior a la "/" y pasar a número.
    for col in (13, 14):
        if col < len(df.columns):
            nombre = df.columns[col]
            df[nombre] = (
                df.iloc[:, col].fillna("").astype(str).str.split("/", n=1).str[0].str.strip()
            )
            df[nombre] = pd.to_numeric(df[nombre], errors="coerce")
    log.ok("✓ Columnas GAP y SAE limpiadas")

    log.titulo("📊 PASO 2: EXTRACCIÓN DE COLUMNAS")
    series, nombres = [], []
    for valor in COLUMNAS_A_EXTRAER:
        if isinstance(valor, int):
            pos = valor - 1
            if not 0 <= pos < len(df.columns):
                raise ValueError(
                    f"El archivo no tiene suficientes columnas (se esperaban al menos {pos + 1})."
                )
            series.append(df.iloc[:, pos].copy())
            nombres.append(df.columns[pos])
        else:
            series.append(pd.Series([""] * len(df), index=df.index, dtype=object))
            nombres.append(
                {4: "LIC. MED/LL.T/AUS.", 5: "DEMERITO"}.get(len(nombres), "")
            )

    df_eventos = pd.concat(series, axis=1)
    df_eventos.columns = nombres
    df_eventos = df_eventos.rename(columns={"TIPIFICACION": "TIPO DE EVENTO"})
    log.ok(f"✓ Filas extraídas: {len(df_eventos)}")

    log.titulo("📊 PASO 3: PROCESAMIENTO ESPECIALIZADO")
    partes = [df_eventos]
    conteos = {"Eventos principales": len(df_eventos)}
    for etiqueta, clase in (
        ("Contraventores/Detenidos", ContraventoresProcessor),
        ("Integrados/Destacados", IntegradosProcessor),
        ("Colaboradores", ColaboradoresProcessor),
    ):
        log.info(f"→ Procesando: {etiqueta}")
        try:
            df_parte = clase().process_df(df)
        except Exception as exc:  # una parte que falla no frena a las demás
            log.aviso(f"⚠ No se pudo procesar {etiqueta}: {exc}")
            continue
        partes.append(df_parte)
        conteos[etiqueta] = len(df_parte)
        log.ok(f"✓ {len(df_parte)} eventos")

    log.titulo("📊 PASO 4: ARCHIVO UNIFICADO")
    unificado = ruta.parent / "eventos_unificados.xlsx"
    df_total = pd.concat(partes, ignore_index=True, sort=False)
    with pd.ExcelWriter(unificado, engine="openpyxl") as writer:
        df_total.to_excel(writer, sheet_name="eventos", index=False)
    log.ok(f"✅ Generado: {unificado.name} ({len(df_total)} filas)")
    return unificado, conteos


# ─────────────────────────────────────────────────────────────
#  PASO B (opcional): narración
# ─────────────────────────────────────────────────────────────
@contextmanager
def _sin_dialogos_tk(indice_hoja: int, log: Logger):
    """Puente temporal: NarracionProcessor abre ventanitas de Tkinter.

    Desde un hilo de fondo eso no es seguro, así que durante la llamada las
    reemplazamos: la 'pregunta' devuelve la hoja ya elegida en el panel y los
    errores van al log. (Solución definitiva: darle a NarracionProcessor un
    parámetro `hoja` y quitarle tkinter; ver notas del informe.)
    """
    from unittest import mock

    def _error(titulo, mensaje, **_):
        log.error(f"❌ {titulo}: {mensaje}")

    with (
        mock.patch("complementos.narracion.simpledialog.askinteger", return_value=indice_hoja),
        mock.patch("complementos.narracion.messagebox.showerror", side_effect=_error),
    ):
        yield


def _agregar_narracion(ruta_narr: Path, hoja: str, ruta_cmu: Path, unificado: Path, log: Logger):
    from complementos.narracion import NarracionProcessor

    proc = NarracionProcessor(log_fn=log, datos_dir=DATOS_DIR)
    mes = proc.extraer_mes_de_nombre(ruta_cmu.name)
    with pd.ExcelFile(ruta_narr, engine="openpyxl") as libro:
        indice = libro.sheet_names.index(hoja)
    with _sin_dialogos_tk(indice, log):
        return proc.procesar_narracion_file(str(ruta_narr), mes, str(unificado))


# ─────────────────────────────────────────────────────────────
#  PASO C (opcional): presentismo
# ─────────────────────────────────────────────────────────────
def _agregar_presentismo(ruta_pres: Path, unificado: Path, log: Logger):
    from complementos.presentismo import PresentismoProcessor

    ajustar_dias_presentismo(ruta_pres.name)
    return PresentismoProcessor(log_fn=log).procesar_presentismo_file(str(ruta_pres), str(unificado))


# ─────────────────────────────────────────────────────────────
#  ACCIÓN
# ─────────────────────────────────────────────────────────────
def ejecutar(entradas: Entradas, log: Logger) -> Resultado:
    log.titulo("🚀 PRODUCTIVIDAD GENERAL")
    ruta_cmu = entradas.ruta("cmu")
    unificado, conteos = _generar_unificado(ruta_cmu, log)

    problemas: list[str] = []

    ruta_narr = entradas.opcional("narracion")
    if ruta_narr:
        ok, msg = _agregar_narracion(ruta_narr, str(entradas.hoja("narracion")), ruta_cmu, unificado, log)
        (log.ok if ok else log.error)(msg)
        if not ok:
            problemas.append(f"Narración: {msg}")

    ruta_pres = entradas.opcional("presentismo")
    if ruta_pres:
        ok, msg = _agregar_presentismo(ruta_pres, unificado, log)
        (log.ok if ok else log.error)(msg)
        if not ok:
            problemas.append(f"Presentismo: {msg}")

    total_final = len(pd.read_excel(unificado, engine="openpyxl"))
    metricas: dict[str, str | int] = {**conteos, "Total en unificado": total_final}
    if problemas:
        for p in problemas:
            log.aviso(f"⚠ {p}")
        metricas["Pasos con error"] = len(problemas)

    return Resultado(
        mensaje="Proceso completado" + (" con avisos." if problemas else "."),
        metricas=metricas,
        salidas=[unificado],
    )


HERRAMIENTA = Herramienta(
    id="general",
    nombre="Productividad general",
    corto="General",
    icono="INSIGHTS",
    resumen="Arma el archivo unificado de eventos a partir del CMU del mes.",
    pasos=(
        "Lee el CMU, corta en la primera fila vacía y limpia GAP y SAE (deja solo el "
        "número previo a la '/').",
        "Extrae las columnas FECHA, HORA, OPERADOR, TIPO DE EVENTO, BARRIO, ORIGEN, GAP y SAE.",
        "Genera filas adicionales: contraventores y detenidos, eventos integrados/destacados "
        "y colaboradores (apoyo óptico / en evento).",
        "Une todo en 'eventos_unificados.xlsx' (hoja 'eventos').",
        "OPCIONAL — Narración: suma los eventos con narración completa cruzando SAE "
        "contra el CMU del mes (carpeta DATOS).",
        "OPCIONAL — Presentismo: suma licencias médicas, llegadas tarde y ausentes.",
    ),
    entradas=(
        ArchivoSpec("cmu", "CMU del mes", "Ej: CMU - SEPTIEMBRE.xlsx"),
        ArchivoSpec(
            "narracion",
            "Control Diario (Narración)",
            "Opcional",
            obligatorio=False,
            pide_hoja=True,
        ),
        ArchivoSpec("presentismo", "Presentismo", "Opcional", obligatorio=False),
    ),
    genera="eventos_unificados.xlsx (junto al CMU) + archivos de presentismo en DATOS/.",
    acciones=(Accion("Generar unificado", "PLAY_ARROW", ejecutar),),
)
