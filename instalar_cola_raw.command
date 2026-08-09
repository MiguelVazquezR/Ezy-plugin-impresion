#!/bin/bash
# ============================================================================
# Instalar_Cola_Raw.command — Configura la impresora térmica como cola RAW
# en macOS para que Ezy Plugin pueda enviar comandos ESC/POS directos.
#
# Uso (en la Mac del suscriptor):
#   1. Clic derecho sobre este archivo → Abrir → Abrir (paso de seguridad)
#   2. Escribe tu contraseña cuando la pida
#   3. Elige impresora USB o red y sigue las instrucciones
#
# Resultado: crea la cola "EZY_Ticket", la deja como predeterminada y verifica.
# ============================================================================

set -e

echo "==========================================================="
echo "  Ezy Plugin — Configurador de Impresora (cola RAW)"
echo "==========================================================="
echo ""

# 1) Permisos de administrador (necesarios para lpadmin)
echo "Se necesita tu contraseña de usuario para administrar impresoras."
echo "La escribirás ahora (no se muestra mientras escribes):"
sudo -v

# 2) Detectar impresoras conectadas
echo ""
echo "--- Impresoras detectadas por macOS ---"
lpstat -p -d 2>/dev/null || true
echo ""

echo "¿Cómo quieres registrar la impresora?"
echo "  1) USB  (impresora conectada por cable USB)"
echo "  2) Red / Wi-Fi (impresora con IP)"
echo ""
read -p "Elige 1 o 2 y presiona Enter: " TIPO

URI=""

if [ "$TIPO" = "1" ]; then
    echo ""
    echo "Buscando dispositivos USB disponibles..."
    lpinfo -v 2>/dev/null | grep -i "usb" || echo "(No se encontraron dispositivos USB)"
    echo ""
    echo "Escribe el URI de tu impresora (la línea completa que empieza"
    echo "con usb://, por ejemplo: usb://EPSON/TM-T20?serial=ABC123)."
    read -p "URI: " URI
    if [ -z "$URI" ]; then
        echo "❌ URI vacío. Cancela y vuelve a intentar."
        exit 1
    fi
else
    echo ""
    read -p "Dirección IP de la impresora (ej. 192.168.1.50): " IP
    if [ -z "$IP" ]; then
        echo "❌ IP vacía. Cancela y vuelve a intentar."
        exit 1
    fi
    # Probamos socket (9100, común en impresoras térmicas) o IPP
    if curl -sf --connect-timeout 3 "ipp://$IP/ipp/print" >/dev/null 2>&1; then
        URI="ipp://$IP/ipp/print"
    else
        URI="socket://$IP:9100"
    fi
    echo "Usando URI: $URI"
fi

# 3) Crear la cola RAW
echo ""
echo "Creando cola 'EZY_Ticket' en modo raw..."
sudo lpadmin -p EZY_Ticket -E -v "$URI" -m raw
sudo lpadmin -d EZY_Ticket

# 4) Verificar
echo ""
echo "✅ Impresora configurada correctamente. Estado:"
lpstat -p EZY_Ticket

echo ""
echo "==========================================================="
echo "  LISTO. En los ajustes del plugin Ezy usa:"
echo "    - Nombre de impresora:  EZY_Ticket"
echo "==========================================================="
echo ""
echo "Puedes cerrar esta ventana."
echo ""
read -p "Presiona Enter para cerrar..." _