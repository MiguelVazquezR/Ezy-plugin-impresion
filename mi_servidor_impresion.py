# --- Importaciones Finales ---
import os
import sys
import logging
import threading
import requests

# Importaciones de Flask y Servidor
from flask import Flask, jsonify, request
from flask_cors import CORS
from waitress import serve

# Importaciones de librerías de terceros
from pystray import MenuItem, Icon
from PIL import Image
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
    return jsonify({"ok": True, "version": "1.0.2"}) # Versión actualizada

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
    try:
        carga_util = request.get_json()
        nombre_impresora = carga_util.get('nombreImpresora')
        operaciones = carga_util.get('operaciones', [])
        
        # Ancho máximo en puntos (dots) de la impresora.
        # 576px es el estándar para papel de 80mm.
        # 384px es el estándar para papel de 58mm.
        if carga_util.get('anchoImpresora') == '58mm':
            ancho_max_impresora = 384
        else:
            ancho_max_impresora = 576

        if not nombre_impresora:
            raise ValueError("El campo 'nombreImpresora' es requerido.")
        
        logger.info(f"Petición de impresión para '{nombre_impresora}' con {len(operaciones)} operaciones.")

        impresora_virtual = Dummy()
        
        for op in operaciones:
            nombre_op = op.get('nombre')
            args = op.get('argumentos', [])
            logger.info(f"Procesando operación: {nombre_op}")

            if nombre_op == "EscribirTexto":
                impresora_virtual.text(args[0] if args else "")
            elif nombre_op == "Feed":
                impresora_virtual.text("\n" * (int(args[0]) if args else 1))
            elif nombre_op == "TextoSegunPaginaDeCodigos":
                if len(args) < 3: continue
                impresora_virtual.codepage = args[1] 
                impresora_virtual.text(args[2])
            elif nombre_op == "DescargarImagenDeInternetEImprimir":
                if not args or not args[0]:
                    logger.warning("URL de imagen no proporcionada.")
                    continue
                
                url_imagen = args[0]
                # El segundo argumento es el ancho deseado. Es opcional.
                ancho_deseado = args[1] if len(args) > 1 and args[1] is not None else None
                
                logger.info(f"Descargando imagen desde: {url_imagen}")
                
                try:
                    respuesta = requests.get(url_imagen, stream=True)
                    respuesta.raise_for_status()
                    
                    imagen_pil = Image.open(respuesta.raw)
                    ancho_original, alto_original = imagen_pil.size

                    # Determinar el ancho objetivo inicial.
                    # Si el usuario no especifica un ancho, el objetivo es el ancho original.
                    ancho_objetivo = ancho_deseado if ancho_deseado is not None else ancho_original
                    
                    # El ancho final NUNCA puede superar el ancho máximo de la impresora.
                    ancho_final = min(ancho_objetivo, ancho_max_impresora)
                    
                    # Redimensionar solo si el ancho final es diferente al original.
                    if ancho_final != ancho_original:
                        logger.info(f"Redimensionando imagen. Original: {ancho_original}px, Final: {ancho_final}px.")
                        ratio = alto_original / float(ancho_original)
                        alto_final = int(ancho_final * ratio)
                        
                        # Usamos Image.Resampling.LANCZOS para mayor calidad en la redimensión
                        imagen_a_imprimir = imagen_pil.resize((ancho_final, alto_final), Image.Resampling.LANCZOS)
                    else:
                        logger.info(f"Imprimiendo imagen con su ancho original de {ancho_original}px.")
                        imagen_a_imprimir = imagen_pil

                    impresora_virtual.image(imagen_a_imprimir, impl="bitImageRaster")
                    logger.info("Imagen procesada para impresión.")
                    
                except requests.exceptions.RequestException as e:
                    logger.error(f"Error al descargar la imagen: {e}")
                except Exception as e:
                    logger.error(f"Error al procesar la imagen: {e}")

        payload_final_bytes = impresora_virtual.output
        
        if not payload_final_bytes:
            raise ValueError("No se generaron comandos de impresión.")

        hPrinter = win32print.OpenPrinter(nombre_impresora)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Ticket Plugin Flask", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, payload_final_bytes)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)
        
        logger.info(f"Impresión enviada correctamente a {nombre_impresora}.")
        return jsonify({"ok": True, "message": "Operaciones procesadas e impresas correctamente"})

    except Exception as e:
        logger.error(f"Error durante la impresión: {e}")
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