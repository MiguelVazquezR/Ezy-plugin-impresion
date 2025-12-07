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
from pystray import MenuItem, Icon
from PIL import Image, ImageOps
from escpos.printer import Dummy

# Importaciones de Windows
import win32print

# --- Sistema de Logging (sin cambios) ---
logger = logging.getLogger('MiPluginLogger')
logger.setLevel(logging.INFO)
try:
    app_data_path = os.getenv('APPDATA')
    log_dir = os.path.join(app_data_path, 'EzyPlugin')
    os.makedirs(log_dir, exist_ok=True)
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

# --- Definición de Endpoints con Flask ---

@app.route('/version', methods=['GET'])
def get_version():
    logger.info("Solicitud de versión recibida.")
    return jsonify({"ok": True, "version": "1.1.0"}) # Versión incrementada por funcionalidad de cajón

@app.route('/impresoras', methods=['GET'])
def get_impresoras():
    logger.info("Se solicitaron las impresoras.")
    try:
        impresoras_raw = win32print.EnumPrinters(2)
        nombres_impresoras = [impresora[2] for impresora in impresoras_raw]
        logger.info(f"Encontradas: {nombres_impresoras}")
        return jsonify(nombres_impresoras)
    except Exception as e:
        logger.error(f"Error al obtener la lista de impresoras: {e}")
        return jsonify({"ok": False, "message": str(e)}), 500

@app.route('/imprimir', methods=['POST'])
def post_imprimir():
    hPrinter = None
    try:
        carga_util = request.get_json()
        nombre_impresora = carga_util.get('nombreImpresora')
        operaciones = carga_util.get('operaciones', [])
        
        # --- LÓGICA DE ZONA SEGURA ---
        if carga_util.get('anchoImpresora') == '58mm':
            ancho_canvas = 384 # 48 bytes exactos
            ancho_seguro_max = 300 # 300px de imagen + 42px blancos a cada lado
        else:
            ancho_canvas = 576 # 72 bytes exactos (80mm)
            ancho_seguro_max = 512 # Margen estándar de seguridad para 80mm
            
        if not nombre_impresora:
            raise ValueError("El campo 'nombreImpresora' es requerido.")
        
        logger.info(f"Petición para '{nombre_impresora}'. Canvas: {ancho_canvas}px, Zona Segura: {ancho_seguro_max}px")

        hPrinter = win32print.OpenPrinter(nombre_impresora)
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Ticket Plugin Flask", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)

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
                    win32print.WritePrinter(hPrinter, bytes_texto)
                    buffer_texto = Dummy()

                if not args or not args[0]: continue
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
                        
                        win32print.WritePrinter(hPrinter, bytes_fragmento)
                        time.sleep(0.15) 
                        y_pos += CHUNK_HEIGHT

                    # Safety Feed
                    safety_feed = Dummy()
                    safety_feed.text("\n") 
                    win32print.WritePrinter(hPrinter, safety_feed.output)
                    time.sleep(0.1)

                except Exception as e:
                    logger.error(f"Error procesando imagen: {e}")

        bytes_finales = buffer_texto.output
        if bytes_finales:
            win32print.WritePrinter(hPrinter, bytes_finales)

        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
        win32print.ClosePrinter(hPrinter)
        
        logger.info(f"Impresión finalizada en {nombre_impresora}.")
        return jsonify({"ok": True, "message": "Operaciones enviadas correctamente"})

    except Exception as e:
        logger.error(f"Error crítico durante impresión: {e}")
        try:
            if hPrinter:
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                win32print.ClosePrinter(hPrinter)
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
    
    try:
        image = Image.open(resource_path("icon.png"))
    except FileNotFoundError:
        image = Image.new('RGB', (64, 64), 'black')
        logger.warning("No se encontró 'icon.png'. Usando ícono por defecto.")

    menu = (MenuItem('Salir', exit_action),)
    icon = Icon("TuPluginImpresion", image, "Ezy Plugin de Impresión", menu)
    
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    
    logger.info("Iniciando ícono en la bandeja del sistema.")
    icon.run()