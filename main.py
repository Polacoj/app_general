"""
Panel general — Informe Scoring (Flet).

Ejecutar desde la carpeta raíz del proyecto:
    uv run python -m app_general.main

Estructura de la pantalla:
    ┌────────┬───────────────────────────────┬──────────────────┐
    │ Menú   │  Informe de la herramienta    │  Log en vivo     │
    │ lateral│  + archivos + botones         │  (costado)       │
    │        │  + métricas del resultado     │                  │
    └────────┴───────────────────────────────┴──────────────────┘

CONCEPTOS CLAVE:
  • async / await: Flet maneja los eventos con una "autopista" asíncrona. Un handler
    `async def` puede ESPERAR (ej. a que elijas un archivo) sin congelar la ventana.
  • asyncio.to_thread: el trabajo pesado (leer Excel con pandas) corre en un hilo
    aparte; así la interfaz sigue viva y el log avanza en tiempo real.
  • Una sola tarea "bombeadora" lee los Logger de todas las herramientas cada 150 ms y
    dibuja las líneas nuevas. Es el único código que toca los controles del log.
"""

from __future__ import annotations

import asyncio
import os
import subprocess
import sys
import time
import traceback
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:  # permite `python app_general/main.py` y `-m app_general.main`
    sys.path.insert(0, str(RAIZ))

import flet as ft  # noqa: E402
import pandas as pd  # noqa: E402

from app_general.herramientas import HERRAMIENTAS  # noqa: E402
from app_general.nucleo import (  # noqa: E402
    Accion,
    ArchivoSpec,
    Entradas,
    Herramienta,
    LineaLog,
    Logger,
    ProcesoIncompleto,
    Resultado,
)

COLOR_PRINCIPAL = ft.Colors.INDIGO
COLORES_LOG = {
    "info": "#d4d4d4",
    "ok": "#6ee7a8",
    "aviso": "#fbbf24",
    "error": "#f87171",
    "titulo": "#7dd3fc",
}
MAX_LINEAS_LOG = 2000


def icono(nombre: str):
    """Convierte 'GAVEL' en ft.Icons.GAVEL (con ícono de respaldo si no existe)."""
    return getattr(ft.Icons, nombre, ft.Icons.CIRCLE)


def abrir_carpeta(ruta: Path) -> None:
    """Abre la carpeta en el explorador del sistema (Mac, Windows o Linux)."""
    if sys.platform == "darwin":
        subprocess.Popen(["open", str(ruta)])
    elif sys.platform == "win32":
        os.startfile(ruta)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", str(ruta)])


# ═════════════════════════════════════════════════════════════
#  VISTA DE UNA HERRAMIENTA
# ═════════════════════════════════════════════════════════════
class VistaHerramienta:
    """Construye y controla la pantalla completa de UNA herramienta."""

    def __init__(self, app: "App", herramienta: Herramienta) -> None:
        self.app = app
        self.h = herramienta
        self.log = Logger()

        self.rutas: dict[str, Path] = {}
        self.campos: dict[str, ft.TextField] = {}
        self.selectores_hoja: dict[str, ft.Dropdown] = {}
        self.botones: list[ft.Button] = []
        self.ultima_carpeta: Path | None = None

        self.anillo = ft.ProgressRing(visible=False, width=22, height=22)
        self.estado = ft.Text("Listo", size=12, color=ft.Colors.GREY_700)
        self.metricas = ft.Row(wrap=True, spacing=12, run_spacing=12)
        self.salidas_texto = ft.Text("", size=12, selectable=True)
        self.tablas = ft.Column(spacing=18)
        self.btn_abrir = ft.TextButton(
            "Abrir carpeta de salida",
            icon=ft.Icons.FOLDER_OPEN,
            visible=False,
            on_click=self._abrir_salida,
        )
        self.lista_log = ft.ListView(expand=True, auto_scroll=True, spacing=2)

        self.control = self._construir()

    # ── construcción ─────────────────────────────────────────
    def _construir(self) -> ft.Control:
        h = self.h

        encabezado = ft.Row(
            [
                ft.Icon(icono(h.icono), size=38, color=COLOR_PRINCIPAL),
                ft.Column(
                    [
                        ft.Text(h.nombre, size=26, weight=ft.FontWeight.BOLD),
                        ft.Text(h.resumen, size=13, color=ft.Colors.GREY_700),
                    ],
                    spacing=2,
                    expand=True,
                ),
            ],
            spacing=14,
        )

        informe = ft.Card(
            content=ft.Container(
                padding=16,
                content=ft.Column(
                    [
                        ft.Text("¿Qué hace esta herramienta?", weight=ft.FontWeight.BOLD),
                        *[
                            ft.Text(f"{i}.  {paso}", size=13, selectable=True)
                            for i, paso in enumerate(h.pasos, start=1)
                        ],
                        ft.Divider(height=8),
                        ft.Text(f"📄 Genera: {h.genera}", size=12, color=ft.Colors.GREY_800),
                    ],
                    spacing=6,
                ),
            )
        )

        filas = [self._fila_archivo(spec) for spec in h.entradas]

        botones = []
        for accion in h.acciones:
            boton = ft.Button(
                accion.etiqueta,
                icon=icono(accion.icono),
                on_click=self._crear_handler(accion),
            )
            self.botones.append(boton)
            botones.append(boton)

        centro = ft.Column(
            [
                encabezado,
                informe,
                ft.Text("Archivos de entrada", weight=ft.FontWeight.BOLD),
                *filas,
                ft.Row([*botones, self.anillo, self.estado], spacing=12),
                self.metricas,
                self.salidas_texto,
                self.btn_abrir,
                self.tablas,
            ],
            spacing=14,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

        panel_log = ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Log en vivo", weight=ft.FontWeight.BOLD, color="#e5e7eb", expand=True),
                        ft.IconButton(
                            icon=ft.Icons.DELETE_SWEEP,
                            icon_color="#9ca3af",
                            tooltip="Limpiar log",
                            on_click=self._limpiar_log,
                        ),
                    ]
                ),
                ft.Container(content=self.lista_log, expand=True),
            ],
            expand=True,
        )

        return ft.Row(
            [
                ft.Container(content=centro, expand=3, padding=20),
                ft.VerticalDivider(width=1),
                ft.Container(content=panel_log, expand=2, padding=12, bgcolor="#1e1e1e"),
            ],
            expand=True,
            spacing=0,
        )

    def _fila_archivo(self, spec: ArchivoSpec) -> ft.Control:
        etiqueta = spec.etiqueta + ("" if spec.obligatorio else " (opcional)")
        campo = ft.TextField(
            label=etiqueta,
            hint_text=spec.ayuda or "Sin seleccionar",
            read_only=True,
            expand=True,
            dense=True,
        )
        self.campos[spec.clave] = campo

        async def al_examinar(e):
            await self._elegir_archivo(spec)

        def al_quitar(e):
            self._quitar_archivo(spec)

        fila = ft.Row(
            [
                campo,
                ft.IconButton(icon=ft.Icons.CLOSE, tooltip="Quitar", on_click=al_quitar),
                ft.Button("Examinar", icon=ft.Icons.FOLDER_OPEN, on_click=al_examinar),
            ]
        )
        if not spec.pide_hoja:
            return fila

        selector = ft.Dropdown(label="Hoja", width=320, visible=False, options=[])
        self.selectores_hoja[spec.clave] = selector
        return ft.Column([fila, selector], spacing=6)

    def _crear_handler(self, accion: Accion):
        async def al_click(e):
            await self._ejecutar(accion)

        return al_click

    # ── selección de archivos ────────────────────────────────
    async def _elegir_archivo(self, spec: ArchivoSpec) -> None:
        archivos = await self.app.picker.pick_files(
            dialog_title=f"Seleccionar: {spec.etiqueta}",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx", "xlsm"],
            allow_multiple=False,
        )
        if not archivos or not archivos[0].path:
            return
        ruta = Path(archivos[0].path)
        self.rutas[spec.clave] = ruta
        self.campos[spec.clave].value = str(ruta)
        self.log.info(f"Archivo seleccionado — {spec.etiqueta}: {ruta.name}")
        if spec.pide_hoja:
            self._cargar_hojas(spec, ruta)
        self.app.page.update()

    def _cargar_hojas(self, spec: ArchivoSpec, ruta: Path) -> None:
        selector = self.selectores_hoja[spec.clave]
        try:
            with pd.ExcelFile(ruta) as libro:
                hojas = list(libro.sheet_names)
        except Exception as exc:
            self.log.error(f"❌ No se pudo leer las hojas de {ruta.name}: {exc}")
            return
        selector.options = [ft.DropdownOption(key=h, text=h) for h in hojas]
        selector.value = hojas[0]
        selector.visible = len(hojas) > 1  # con una sola hoja no hace falta preguntar
        if len(hojas) == 1:
            self.log.info(f"Hoja única seleccionada automáticamente: {hojas[0]}")
        else:
            self.log.info(f"El libro tiene {len(hojas)} hojas: elegí una en el selector.")

    def _quitar_archivo(self, spec: ArchivoSpec) -> None:
        self.rutas.pop(spec.clave, None)
        self.campos[spec.clave].value = ""
        selector = self.selectores_hoja.get(spec.clave)
        if selector:
            selector.visible = False
            selector.options = []
            selector.value = None
        self.app.page.update()

    def _leer_entradas(self) -> Entradas:
        faltan = [s.etiqueta for s in self.h.entradas if s.obligatorio and s.clave not in self.rutas]
        if faltan:
            raise ValueError("Falta seleccionar: " + ", ".join(faltan))
        hojas = {
            clave: selector.value
            for clave, selector in self.selectores_hoja.items()
            if clave in self.rutas and selector.value
        }
        return Entradas(archivos=dict(self.rutas), hojas=hojas)

    # ── ejecución ────────────────────────────────────────────
    def _poner_estado(self, texto: str, color: str) -> None:
        self.estado.value = texto
        self.estado.color = color

    def _ocupado(self, valor: bool) -> None:
        self.anillo.visible = valor
        for boton in self.botones:
            boton.disabled = valor

    async def _ejecutar(self, accion: Accion) -> None:
        try:
            entradas = self._leer_entradas()
        except ValueError as exc:
            self.log.aviso(f"⚠ {exc}")
            self._poner_estado("Faltan datos", ft.Colors.ORANGE_800)
            self.app.page.update()
            return

        self._ocupado(True)
        self._poner_estado("Procesando…", COLOR_PRINCIPAL)
        self.metricas.controls.clear()
        self.tablas.controls.clear()
        self.salidas_texto.value = ""
        self.btn_abrir.visible = False
        self.app.page.update()

        self.log.titulo(f"▶ {accion.etiqueta.upper()}")
        inicio = time.perf_counter()
        try:
            # El trabajo pesado va a un hilo; mientras tanto la ventana sigue respondiendo.
            resultado = await asyncio.to_thread(accion.ejecutar, entradas, self.log)
        except ProcesoIncompleto as exc:
            self.log.aviso(f"⚠ {exc}")
            self._poner_estado("Incompleto: revisá el log", ft.Colors.ORANGE_800)
        except Exception as exc:
            self.log.error(f"❌ {exc}")
            self.log.info(traceback.format_exc())
            self._poner_estado("Error: revisá el log", ft.Colors.RED_700)
        else:
            segundos = time.perf_counter() - inicio
            self.log.ok(f"✅ {resultado.mensaje} ({segundos:.1f} s)")
            self._poner_estado("Completado", ft.Colors.GREEN_700)
            self._mostrar_resultado(resultado)
        finally:
            self._ocupado(False)
            self.app.page.update()

    def _mostrar_resultado(self, res: Resultado) -> None:
        self.metricas.controls = [self._tarjeta(k, v) for k, v in res.metricas.items()]
        self.tablas.controls = [self._tabla(t, df) for t, df in res.tablas.items()]
        if res.salidas:
            self.salidas_texto.value = "Archivos generados:\n" + "\n".join(f"• {p}" for p in res.salidas)
            self.ultima_carpeta = res.salidas[0].parent
            self.btn_abrir.visible = True

    @staticmethod
    def _tarjeta(etiqueta: str, valor: object) -> ft.Control:
        return ft.Container(
            width=200,
            padding=14,
            border_radius=10,
            bgcolor=ft.Colors.INDIGO_50,
            content=ft.Column(
                [
                    ft.Text(str(valor), size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.INDIGO_900),
                    ft.Text(etiqueta, size=12, color=ft.Colors.GREY_700),
                ],
                spacing=2,
            ),
        )

    @staticmethod
    def _tabla(titulo: str, df: pd.DataFrame) -> ft.Control:
        """Dibuja un DataFrame como tabla; la fila 'Total…' va en negrita."""
        columnas = [
            ft.DataColumn(
                ft.Text(str(c), weight=ft.FontWeight.BOLD),
                numeric=pd.api.types.is_numeric_dtype(df[c]),
            )
            for c in df.columns
        ]
        filas = []
        for fila in df.itertuples(index=False):
            es_total = str(fila[0]).lower().startswith("total")
            negrita = ft.FontWeight.BOLD if es_total else None
            filas.append(
                ft.DataRow(cells=[ft.DataCell(ft.Text(str(v), weight=negrita)) for v in fila])
            )
        return ft.Column(
            [
                ft.Text(titulo, weight=ft.FontWeight.BOLD),
                ft.Row(
                    [
                        ft.DataTable(
                            columns=columnas,
                            rows=filas,
                            heading_row_color=ft.Colors.INDIGO_50,
                            column_spacing=40,
                        )
                    ],
                    scroll=ft.ScrollMode.AUTO,
                ),
            ],
            spacing=8,
        )

    def _abrir_salida(self, e) -> None:
        if self.ultima_carpeta:
            abrir_carpeta(self.ultima_carpeta)

    # ── log ──────────────────────────────────────────────────
    def drenar_log(self) -> bool:
        """Dibuja los mensajes nuevos. Devuelve True si hubo algo para mostrar."""
        lineas: list[LineaLog] = self.log.vaciar()
        for linea in lineas:
            self.lista_log.controls.append(
                ft.Text(
                    f"{linea.hora}  {linea.texto}",
                    size=12,
                    font_family="Courier New",
                    color=COLORES_LOG.get(linea.nivel, COLORES_LOG["info"]),
                    weight=ft.FontWeight.BOLD if linea.nivel == "titulo" else None,
                    selectable=True,
                )
            )
        exceso = len(self.lista_log.controls) - MAX_LINEAS_LOG
        if exceso > 0:
            del self.lista_log.controls[:exceso]
        return bool(lineas)

    def _limpiar_log(self, e) -> None:
        self.lista_log.controls.clear()
        self.app.page.update()


# ═════════════════════════════════════════════════════════════
#  APLICACIÓN
# ═════════════════════════════════════════════════════════════
class App:
    def __init__(self, page: ft.Page) -> None:
        self.page = page
        self.picker = ft.FilePicker()
        self.vistas = [VistaHerramienta(self, h) for h in HERRAMIENTAS]
        self.contenido = ft.Container(expand=True)
        self._tarea_log: asyncio.Task | None = None

        destinos = [
            ft.NavigationRailDestination(
                icon=ft.Icons.DASHBOARD_OUTLINED,
                selected_icon=ft.Icons.DASHBOARD,
                label="Inicio",
            )
        ]
        destinos += [
            ft.NavigationRailDestination(icon=icono(h.icono), label=h.corto) for h in HERRAMIENTAS
        ]
        self.rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=96,
            destinations=destinos,
            on_change=self._al_navegar,
        )
        self.inicio = self._construir_inicio()

    def construir(self) -> None:
        page = self.page
        page.title = "Informe Scoring — Panel general"
        page.theme_mode = ft.ThemeMode.LIGHT
        page.theme = ft.Theme(color_scheme_seed=COLOR_PRINCIPAL)
        page.padding = 0
        page.window.width = 1440
        page.window.height = 860
        page.services.append(self.picker)

        self.contenido.content = self.inicio
        page.add(ft.Row([self.rail, ft.VerticalDivider(width=1), self.contenido], expand=True, spacing=0))
        # Tarea de fondo que dibuja el log en vivo.
        self._tarea_log = asyncio.create_task(self._bombear_logs())

    # ── navegación ───────────────────────────────────────────
    def _al_navegar(self, e) -> None:
        self._ir_a(e.control.selected_index)

    def _ir_a(self, indice: int) -> None:
        self.rail.selected_index = indice
        self.contenido.content = self.inicio if indice == 0 else self.vistas[indice - 1].control
        self.page.update()

    def _construir_inicio(self) -> ft.Control:
        tarjetas = []
        for i, h in enumerate(HERRAMIENTAS, start=1):
            tarjetas.append(
                ft.Container(
                    width=330,
                    padding=18,
                    border_radius=12,
                    bgcolor=ft.Colors.INDIGO_50,
                    ink=True,
                    on_click=lambda e, indice=i: self._ir_a(indice),
                    content=ft.Column(
                        [
                            ft.Icon(icono(h.icono), size=32, color=COLOR_PRINCIPAL),
                            ft.Text(h.nombre, size=17, weight=ft.FontWeight.BOLD),
                            ft.Text(h.resumen, size=12, color=ft.Colors.GREY_800),
                        ],
                        spacing=6,
                    ),
                )
            )
        return ft.Column(
            [
                ft.Text("Panel general de productividad", size=28, weight=ft.FontWeight.BOLD),
                ft.Text(
                    "Elegí una herramienta: verás qué hace, cargás los archivos y seguís el "
                    "proceso en el log en vivo.",
                    color=ft.Colors.GREY_700,
                ),
                ft.Row(tarjetas, wrap=True, spacing=16, run_spacing=16),
            ],
            spacing=18,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

    # ── log en vivo ──────────────────────────────────────────
    async def _bombear_logs(self) -> None:
        while True:
            hubo_cambios = False
            for vista in self.vistas:
                hubo_cambios |= vista.drenar_log()
            if hubo_cambios:
                self.page.update()
            await asyncio.sleep(0.15)


async def main(page: ft.Page) -> None:
    App(page).construir()


if __name__ == "__main__":
    ft.run(main)
