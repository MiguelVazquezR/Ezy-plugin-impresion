# --- Importaciones Finales ---
import os
import sys
import logging
import threading
import requests
import time

# Importaciones de Flask y Servidor
from flask import Flask, jsonify, request
from flask_cors import CORS
from waitress import serve

# Importaciones de librerías de terceros
from PIL import Image, ImageOps
from escpos.printer import Dummy

# Capa de abstracción de impresión (Windows: win32print / macOS: CUPS / DryRun)
from backend_impresora import crear_backend, obtener_directorio_plugins

# Ícono de bandeja opcional: si pystray (o PyObjC en macOS) no está disponible,
# el servidor sigue funcionando igual, solo sin ícono en la barra de menú.
try:
    from pystray import MenuItem, Icon
except Exception:
    MenuItem = None
    Icon = None

# --- Sistema de Logging multiplataforma ---
# Windows: %APPDATA%\EzyPlugin\impresion.log
# macOS:   ~/Library/Application Support/EzyPlugin/impresion.log
logger = logging.getLogger('MiPluginLogger')
logger.setLevel(logging.INFO)
try:
    log_dir = obtener_directorio_plugins()
    log_file_path = os.path.join(log_dir, 'impresion.log')
    file_handler = logging.FileHandler(log_file_path)
    file_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.info("Sistema de logging inicializado correctamente.")
except Exception as e:
    logger.addHandler(logging.NullHandler())

# --- Configuración de la Aplicación Flask ---
app = Flask(__name__)
CORS(app)


def resolver_backend():
    """
    Devuelve el backend de impresión según el entorno:
      - EZYPLUGIN_DRYRUN=1  → escribe los bytes a un .bin (pruebas sin impresora)
      - Windows             → win32print
      - macOS               → CUPS nativo (lp / lpstat)
    """
    modo = 'dryrun' if os.environ.get('EZYPLUGIN_DRYRUN') == '1' else 'auto'
    return crear_backend(modo)


# --- Definición de Endpoints con Flask ---

@app.route('/version', methods=['GET'])
def get_version():
    logger.info("Solicitud de versión recibida.")
    return jsonify({"ok": True, "version": "1.2.0"})  # Versión: soporte macOS + cajón


@app.route('/impresoras', methods=['GET'])
def get_impresoras():
    logger.info("Se solicitaron las impresoras.")
    try:
        backend = resolver_backend()
        nombres_impresoras = backend.listar_impresoras()
        logger.info(f"Encontradas: {nombres_impresoras}")
        return jsonify(nombres_impresoras)
    except Exception as e:
        logger.error(f"Error al obtener la lista de impresoras: {e}")
        return jsonify({"ok": False, "message": str(e)}), 500


@app.route('/imprimir', methods=['POST'])
def post_imprimir():
    backend = None
    try:
        carga_util = request.get_json()
        nombre_impresora = carga_util.get('nombreImpresora')
        operaciones = carga_util.get('operaciones', [])

        # --- LÓGICA DE ZONA SEGURA ---
        if carga_util.get('anchoImpresora') == '58mm':
            ancho_canvas = 384  # 48 bytes exactos
            ancho_seguro_max = 300  # 300px de imagen + 42px blancos a cada lado
        else:
            ancho_canvas = 576  # 72 bytes exactos (80mm)
            ancho_seguro_max = 512  # Margen estándar de seguridad para 80mm

        if not nombre_impresora:
            raise ValueError("El campo 'nombreImpresora' es requerido.")

        logger.info(f"Petición para '{nombre_impresora}'. Canvas: {ancho_canvas}px, Zona Segura: {ancho_seguro_max}px")

        backend = resolver_backend()
        backend.abrir(nombre_impresora)

        buffer_texto = Dummy()

        for op in operaciones:
            nombre_op = op.get('nombre')
            args = op.get('argumentos', [])
            logger.info(f"Procesando operación: {nombre_op}")

            if nombre_op == "EscribirTexto":
                buffer_texto.text(args[0] if args else "")

            elif nombre_op == "AbrirCajon":
                # La mayoría de impresoras usan el Pin 2 (Standard).
                # Enviamos el comando y procesamos inmediatamente para que el cajón abra YA.
                try:
                    # Pin 2 (Standard)
                    buffer_texto.cashdraw(2)
                    # Opcional: Pin 5 (algunas impresoras raras lo usan, no suele hacer daño enviar ambos)
                    # buffer_texto.cashdraw(5)
                except Exception as e:
                    logger.error(f"Error generando comando cajón: {e}")

            elif nombre_op == "Feed":
                buffer_texto.text("\n" * (int(args[0]) if args else 1))

            elif nombre_op == "TextoSegunPaginaDeCodigos":
                if len(args) >= 3:
                    try:
                        buffer_texto.codepage = args[1]
                    except:
                        pass
                    buffer_texto.text(args[2])

            elif nombre_op == "DescargarImagenDeInternetEImprimir":
                # Vaciar buffer de texto previo (incluyendo comando de cajón si lo hubiera)
                bytes_texto = buffer_texto.output
                if bytes_texto:
                    backend.escribir(bytes_texto)
                    buffer_texto = Dummy()

                if not args or not args[0]:
                    continue
                url_imagen = args[0]
                ancho_deseado = args[1] if len(args) > 1 and args[1] is not None else None

                try:
                    logger.info(f"Descargando imagen: {url_imagen}")
                    respuesta = requests.get(url_imagen, stream=True, timeout=20)
                    respuesta.raise_for_status()

                    # 1. Cargar y Sanear
                    imagen_pil = Image.open(respuesta.raw).convert("RGBA")
                    fondo_blanco = Image.new("RGB", imagen_pil.size, (255, 255, 255))
                    fondo_blanco.paste(imagen_pil, mask=imagen_pil.split()[3])
                    imagen_saneada = fondo_blanco

                    # 2. Redimensionar respetando la ZONA SEGURA
                    ancho_original, alto_original = imagen_saneada.size

                    if ancho_deseado is not None:
                        target = int(ancho_deseado)
                    else:
                        target = ancho_original

                    ancho_final_contenido = min(target, ancho_seguro_max)

                    ratio = alto_original / float(ancho_original)
                    alto_final_contenido = int(ancho_final_contenido * ratio)

                    imagen_redimensionada = imagen_saneada.resize((ancho_final_contenido, alto_final_contenido), Image.Resampling.LANCZOS)

                    # 3. Canvas (Padding Blanco)
                    imagen_canvas = Image.new("RGB", (ancho_canvas, alto_final_contenido), (255, 255, 255))
                    pos_x = (ancho_canvas - ancho_final_contenido) // 2
                    imagen_canvas.paste(imagen_redimensionada, (pos_x, 0))

                    # 4. Dither y Streaming
                    imagen_final = imagen_canvas.convert('1', dither=Image.Dither.FLOYDSTEINBERG)

                    CHUNK_HEIGHT = 60
                    y_pos = 0

                    logger.info(f"Enviando contenido de {ancho_final_contenido}px centrado en canvas de {ancho_canvas}px...")

                    while y_pos < alto_final_contenido:
                        bottom = min(y_pos + CHUNK_HEIGHT, alto_final_contenido)
                        box = (0, y_pos, ancho_canvas, bottom)
                        fragmento = imagen_final.crop(box)

                        chunk_d = Dummy()
                        chunk_d.image(fragmento, impl="bitImageRaster")
                        bytes_fragmento = chunk_d.output

                        backend.escribir(bytes_fragmento)
                        time.sleep(0.15)
                        y_pos += CHUNK_HEIGHT

                    # Safety Feed
                    safety_feed = Dummy()
                    safety_feed.text("\n")
                    backend.escribir(safety_feed.output)
                    time.sleep(0.1)

                except Exception as e:
                    logger.error(f"Error procesando imagen: {e}")

        bytes_finales = buffer_texto.output
        if bytes_finales:
            backend.escribir(bytes_finales)

        backend.cerrar()

        logger.info(f"Impresión finalizada en {nombre_impresora}.")
        return jsonify({"ok": True, "message": "Operaciones enviadas correctamente"})

    except Exception as e:
        logger.error(f"Error crítico durante impresión: {e}")
        try:
            if backend:
                backend.cancelar()
        except:
            pass
        return jsonify({"ok": False, "message": str(e)}), 500


def run_server():
    logger.info("Iniciando servidor Waitress en el puerto 8000.")
    serve(app, host='127.0.0.1', port=8000)


def exit_action(icon, item):
    logger.info("Petición de salida recibida. Deteniendo...")
    icon.stop()


if __name__ == '__main__':
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    if Icon is None:
        # Sin pystray/PyObjC: servidor queda activo sin ícono de bandeja.
        logger.info("pystray no disponible; ejecutando sin ícono de bandeja.")
        try:
            server_thread.join()
        except KeyboardInterrupt:
            logger.info("Detenido por el usuario.")
    else:
        try:
            image = Image.open(resource_path("icon.png"))
        except FileNotFoundError:
            image = Image.new('RGB', (64, 64), 'black')
            logger.warning("No se encontró 'icon.png'. Usando ícono por defecto.")

        menu = (MenuItem('Salir', exit_action),)
        icon = Icon("TuPluginImpresion", image, "Ezy Plugin de Impresión", menu)

        logger.info("Iniciando ícono en la bandeja del sistema.")
        try:
            icon.run()
        except Exception as e:
            # pystray puede fallar en sesiones sin interfaz gráfica (SSH, CI,
            # runners sin WindowServer). El servidor sigue vivo y funcional.
            logger.warning(f"No se pudo mostrar el ícono de la bandeja: {e}")
            logger.info("Manteniendo el servidor activo sin ícono de bandeja.")
            server_thread.join()
