# -*- coding: utf-8 -*-
"""
backend_impresora.py — Capa de abstracción de impresión multiplataforma.

Windows  → usa win32print (RAW, comportamiento idéntico al plugin actual).
macOS    → usa CUPS nativo vía subprocess (`lpstat` para listar y `lp -o raw`
           para enviar bytes ESC/POS sin filtros).
DryRun   → escribe los bytes generados a un archivo .bin (para pruebas sin
           impresora física, útil en Windows o en CI de GitHub Actions).

Los bytes ESC/POS son generados por python-escpos en memoria y son idénticos
en todos los sistemas; solo cambia la forma de entregarlos a la impresora.
"""

import os
import sys
import time
import logging
import subprocess
import tempfile
from io import BytesIO

logger = logging.getLogger('MiPluginLogger')


def obtener_directorio_plugins():
    """Ruta estándar de datos del plugin según el sistema operativo."""
    if sys.platform == 'darwin':
        base = os.path.join(os.path.expanduser('~'),
                            'Library', 'Application Support', 'EzyPlugin')
    elif sys.platform == 'win32':
        base = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')),
                            'EzyPlugin')
    else:
        base = os.path.join(os.path.expanduser('~'), '.ezyplugin')
    os.makedirs(base, exist_ok=True)
    return base


class BackendWindows:
    """Entrega bytes ESC/POS a la impresora vía win32print (solo Windows)."""

    def __init__(self):
        # Import dinámico: solo se ejecuta en Windows. PyInstaller lo detecta
        # por bytecode y lo incluye en el binario de Windows sin problema.
        import win32print
        self._win32print = win32print
        self._hPrinter = None
        self._abierto = False

    def listar_impresoras(self):
        impresoras_raw = self._win32print.EnumPrinters(2)
        return [impresora[2] for impresora in impresoras_raw]

    def abrir(self, nombre_impresora):
        self._hPrinter = self._win32print.OpenPrinter(nombre_impresora)
        self._win32print.StartDocPrinter(self._hPrinter, 1,
                                         ("Ticket Plugin Flask", None, "RAW"))
        self._win32print.StartPagePrinter(self._hPrinter)
        self._abierto = True

    def escribir(self, datos):
        if not self._abierto:
            raise RuntimeError("Debe llamar a abrir() antes de escribir.")
        self._win32print.WritePrinter(self._hPrinter, datos)

    def cerrar(self):
        if self._abierto:
            self._win32print.EndPagePrinter(self._hPrinter)
            self._win32print.EndDocPrinter(self._hPrinter)
            self._win32print.ClosePrinter(self._hPrinter)
            self._abierto = False

    def cancelar(self):
        # En Windows hay que cerrar el handle sí o sí para no dejarlo colgado.
        self.cerrar()


class BackendMacOS:
    """
    Entrega bytes ESC/POS usando CUPS nativo de macOS.

    En macOS, CUPS filtra los trabajos por defecto (espera PDF/PostScript).
    Para que acepte comandos ESC/POS raw es necesario que la impresora esté
    registrada como cola RAW:

        sudo lpadmin -p EZY_Ticket -E -v <uri> -m raw

    (Ver GUIA_MAC.md / instalar_cola_raw.command incluidos en el paquete.)
    """

    def __init__(self):
        self._nombre_impresora = None
        self._buffer = BytesIO()

    def listar_impresoras(self):
        try:
            resultado = subprocess.run(["lpstat", "-p"],
                                       capture_output=True, text=True,
                                       timeout=10)
        except Exception as e:
            logger.error(f"Error ejecutando lpstat: {e}")
            return []
        nombres = []
        for linea in resultado.stdout.splitlines():
            partes = linea.split()
            if len(partes) >= 2 and partes[0] == "printer":
                nombres.append(partes[1])
        return nombres

    def abrir(self, nombre_impresora):
        self._nombre_impresora = nombre_impresora
        self._buffer = BytesIO()

    def escribir(self, datos):
        self._buffer.write(datos)

    def cerrar(self):
        if not self._nombre_impresora:
            raise RuntimeError("Debe llamar a abrir() antes de cerrar.")
        bytes_totales = self._buffer.getvalue()
        if not bytes_totales:
            logger.info("Nada que imprimir; no se envía trabajo a CUPS.")
            return

        ruta_temporal = None
        try:
            archivo_temporal = tempfile.NamedTemporaryFile(
                prefix="ezyplugin_", suffix=".bin", delete=False)
            ruta_temporal = archivo_temporal.name
            archivo_temporal.write(bytes_totales)
            archivo_temporal.close()

            logger.info(
                f"Enviando {len(bytes_totales)} bytes a CUPS "
                f"(cola '{self._nombre_impresora}') en modo raw...")

            proceso = subprocess.run(
                ["lp", "-d", self._nombre_impresora,
                 "-t", "EzyPlugin Ticket", "-o", "raw", ruta_temporal],
                capture_output=True, text=True, timeout=60)

            if proceso.returncode != 0:
                raise RuntimeError(
                    f"lp falló (código {proceso.returncode}): "
                    f"{proceso.stderr.strip()}")

            logger.info(f"CUPS aceptó el trabajo: {proceso.stdout.strip()}")
        finally:
            if ruta_temporal and os.path.exists(ruta_temporal):
                try:
                    os.remove(ruta_temporal)
                except OSError:
                    pass

    def cancelar(self):
        # No se envía nada si hubo un error a mitad de generación.
        self._nombre_impresora = None
        self._buffer = BytesIO()


class BackendDryRun:
    """
    Escribe los bytes ESC/POS a un archivo .bin en la carpeta de datos del
    plugin. Sirve para probar el pipeline completo (texto, feed, cajón,
    imágenes, dithering) sin necesidad de impresora física.

    Se activa con la variable de entorno EZYPLUGIN_DRYRUN=1.
    """

    def __init__(self):
        self._real = _crear_backend_nativo()
        self._nombre_impresora = None
        self._buffer = BytesIO()

    def listar_impresoras(self):
        return self._real.listar_impresoras()

    def abrir(self, nombre_impresora):
        self._nombre_impresora = nombre_impresora
        self._buffer = BytesIO()

    def escribir(self, datos):
        self._buffer.write(datos)

    def cerrar(self):
        bytes_totales = self._buffer.getvalue()
        directorio = obtener_directorio_plugins()
        nombre_archivo = (f"impresion_dryrun_"
                          f"{time.strftime('%Y%m%d_%H%M%S')}.bin")
        ruta = os.path.join(directorio, nombre_archivo)
        with open(ruta, "wb") as f:
            f.write(bytes_totales)
        logger.info(
            f"[DRY-RUN] Impresora '{self._nombre_impresora}': "
            f"{len(bytes_totales)} bytes guardados en {ruta}")

    def cancelar(self):
        self._nombre_impresora = None
        self._buffer = BytesIO()


def _crear_backend_nativo():
    if sys.platform == "darwin":
        return BackendMacOS()
    elif sys.platform == "win32":
        return BackendWindows()
    else:
        raise RuntimeError(f"Sistema operativo no soportado: {sys.platform}")


def crear_backend(modo="auto"):
    """
    Crea el backend según el modo:
      - 'auto'  → nativo (win32print en Windows, CUPS en macOS)
      - 'dryrun' → escribe a .bin (pruebas sin impresora)
    """
    if modo == "dryrun":
        return BackendDryRun()
    return _crear_backend_nativo()