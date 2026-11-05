"""
nucleo.py — Piezas compartidas por TODAS las herramientas del panel.

CONCEPTOS CLAVE (para ir aprendiendo):

1. dataclass
   Una forma moderna de declarar "clases que solo guardan datos". Python escribe
   por vos el __init__, el __repr__, etc. Con frozen=True el objeto es inmutable.

2. Logger con cola (queue.Queue)
   Las herramientas corren en un HILO de fondo (para no congelar la ventana).
   Un hilo NO debe tocar la interfaz directamente. Entonces el hilo "deja mensajes
   en un buzón" (la cola) y la interfaz, cada ~150 ms, los retira y los dibuja.
   Esto se llama patrón productor/consumidor y es la base de cualquier log en vivo.

3. Registro de herramientas
   Cada herramienta describe a sí misma con un objeto `Herramienta` (nombre,
   explicación, archivos que pide y qué función ejecutar). La interfaz no sabe nada
   de pandas ni de Excel: solo lee esa descripción. Agregar una herramienta nueva
   = crear un archivo nuevo + una línea en herramientas/__init__.py.
"""

from __future__ import annotations

import queue
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:  # solo para anotaciones; evita importar pandas en el núcleo
    import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
DATOS_DIR = RAIZ / "DATOS"
DATOS_DIR.mkdir(exist_ok=True)


class ProcesoIncompleto(RuntimeError):
    """El proceso se detuvo a propósito (no es un error de programa).

    Ejemplo: la reconciliación de nombres no llegó al 100 % y hay que completar
    el diccionario antes de volver a correr.
    """


# ─────────────────────────────────────────────────────────────
#  LOG
# ─────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class LineaLog:
    nivel: str  # "info" | "ok" | "aviso" | "error" | "titulo"
    texto: str
    hora: str


def _inferir_nivel(texto: str) -> str:
    """Adivina el color de una línea mirando cómo empieza.

    Así podemos reutilizar tus clases viejas (que hacen self.log("✓ ...")) sin
    modificarlas: el emoji del principio decide el nivel.
    """
    t = texto.lstrip()
    if t.startswith(("✓", "✅", "✔")):
        return "ok"
    if t.startswith(("✗", "❌", "ERROR")):
        return "error"
    if t.startswith(("⚠", "AVISO")):
        return "aviso"
    if t.startswith(("═", "─", "📋", "🚀", "📊")):
        return "titulo"
    return "info"


class Logger:
    """Buzón de mensajes seguro para usar desde varios hilos."""

    def __init__(self) -> None:
        self._cola: queue.Queue[LineaLog] = queue.Queue()

    def _emitir(self, nivel: str | None, texto: object) -> None:
        hora = time.strftime("%H:%M:%S")
        for linea in str(texto).splitlines() or [""]:
            self._cola.put(LineaLog(nivel or _inferir_nivel(linea), linea, hora))

    def info(self, texto: object) -> None:
        self._emitir("info", texto)

    def ok(self, texto: object) -> None:
        self._emitir("ok", texto)

    def aviso(self, texto: object) -> None:
        self._emitir("aviso", texto)

    def error(self, texto: object) -> None:
        self._emitir("error", texto)

    def titulo(self, texto: object) -> None:
        self._emitir("titulo", texto)

    def __call__(self, texto: object) -> None:
        """Permite pasar el logger como `log_fn` a tus clases Processor existentes."""
        self._emitir(None, texto)

    def vaciar(self) -> list[LineaLog]:
        """Retira TODOS los mensajes pendientes (lo llama la interfaz)."""
        lineas: list[LineaLog] = []
        while True:
            try:
                lineas.append(self._cola.get_nowait())
            except queue.Empty:
                return lineas


# ─────────────────────────────────────────────────────────────
#  MODELOS
# ─────────────────────────────────────────────────────────────
@dataclass(frozen=True)
class ArchivoSpec:
    """Describe un archivo que la herramienta necesita."""

    clave: str  # identificador interno, ej. "cmu"
    etiqueta: str  # texto visible
    ayuda: str = ""  # pista dentro del campo
    obligatorio: bool = True
    pide_hoja: bool = False  # si True, la UI muestra un selector de hoja


@dataclass
class Entradas:
    """Lo que el usuario eligió en pantalla, listo para la herramienta."""

    archivos: dict[str, Path] = field(default_factory=dict)
    hojas: dict[str, str] = field(default_factory=dict)

    def ruta(self, clave: str) -> Path:
        try:
            return self.archivos[clave]
        except KeyError:
            raise ValueError(f"Falta seleccionar el archivo '{clave}'.") from None

    def opcional(self, clave: str) -> Path | None:
        return self.archivos.get(clave)

    def hoja(self, clave: str) -> str | int:
        """Nombre de la hoja elegida, o 0 (primera hoja) si no se eligió."""
        return self.hojas.get(clave, 0)


@dataclass
class Resultado:
    """Lo que devuelve una herramienta al terminar bien."""

    mensaje: str
    metricas: dict[str, str | int] = field(default_factory=dict)
    salidas: list[Path] = field(default_factory=list)
    # Tablas para mostrar al final del informe: {título: DataFrame}
    tablas: dict[str, "pd.DataFrame"] = field(default_factory=dict)


@dataclass(frozen=True)
class Accion:
    """Un botón de la herramienta (algunas tienen más de uno)."""

    etiqueta: str
    icono: str  # nombre de ft.Icons, como texto
    ejecutar: Callable[[Entradas, Logger], Resultado]


@dataclass(frozen=True)
class Herramienta:
    id: str
    nombre: str
    corto: str  # etiqueta del menú lateral
    icono: str
    resumen: str  # una línea
    pasos: tuple[str, ...]  # "informe aclaratorio": qué hace, paso a paso
    entradas: tuple[ArchivoSpec, ...]
    genera: str  # qué archivos produce
    acciones: tuple[Accion, ...]
