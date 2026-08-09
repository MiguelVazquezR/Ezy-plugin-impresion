# 📱 GUÍA PARA USUARIOS MAC — Plugin de Impresión Ezy Plugin

**Versión del plugin:** 1.2.0
**Compatibilidad:** macOS 11 (Big Sur) o superior — funciona en **Mac Intel** y **Mac Apple Silicon (M1 / M2 / M3 / M4)**

---

## 1. Descargar y descomprimir

1. Recibe el archivo `ezy_plugin_macos.zip` por WhatsApp.
2. Haz **doble clic** sobre el `.zip` para descomprimirlo.
3. Aparecerá la aplicación **`ezy_plugin_macos.app`** (el ícono de Ezy).

> 💡 **Mueve la app a tu carpeta Aplicaciones** (opcional pero recomendado):
> Arrastra `ezy_plugin_macos.app` a la carpeta **Aplicaciones**.

---

## 2. PRIMERA VEZ: abrir la app (paso de seguridad de macOS)

macOS **bloquea por seguridad las apps de desarrolladores no verificados**. Es normal y se resuelve una sola vez:

1. **Clic derecho** sobre `ezy_plugin_macos.app` (o mantén presionada la tecla **Control** y haz clic).
2. En el menú, selecciona **Abrir**.
3. macOS mostrará un aviso: *"ezy_plugin_macos no se puede abrir"* → haz clic en **Abrir** (segundo clic).
4. La app se abrirá. Aparecerá el ícono de la banderita 📌 en la barra superior (junto al reloj).

> Si no aparece la opción "Abrir" o prefieres el método alternativo:
> Ve a **Preferencias del Sistema / Ajustes del Sistema → Privacidad y Seguridad**,
> en la sección **Seguridad**, clic en **"Abrir de todos modos"** junto a la app bloqueada.

✅ A partir de ahora la app se abre con doble clic normal.

---

## 3. IMPORTANTE — Configurar la impresora como "cola RAW" (una sola vez)

En macOS, el sistema filtra los trabajos esperando documentos PDF/PostScript.
Para que la impresora térmica de tickets acepte los comandos directos de Ezy Plugin,
hay que registrarla como **cola RAW**. **Esto se hace UNA sola vez por Mac.**

### Opción A — Con el instalador automático (recomendada)

1. Descomprime el zip y localiza el archivo **`Instalar_Cola_Raw.command`**.
2. **Clic derecho** sobre `Instalar_Cola_Raw.command` → **Abrir** → **Abrir** (mismo paso de seguridad).
3. Se abrirá la **Terminal**, te pedirá:
   - **Tu contraseña de usuario** (la de tu Mac, pedirá `Password:` — escribe y Enter, no se ve mientras escribes).
   - **El nombre de tu impresora** (ej. `TM-T20` — escríbelo tal cual aparece en *Ajustes del Sistema → Impresoras y escáneres*).
4. El script creará la cola `EZY_Ticket`, la dejará como predeterminada y verificará la conexión.
5. Verás `✅ Impresora lista. Escribe "EZY_Ticket" en los ajustes del plugin Ezy.`

### Opción B — Manual (si prefieres)

Abre **Terminal** y ejecuta:

```bash
# 1. Conocer el nombre del dispositivo / URI de tu impresora
lpinfo -v

# 2. Crear la cola RAW (reemplaza el URI por el tuyo)
sudo lpadmin -p EZY_Ticket -E -v "usb://EPSON/TM-T20?serial=..." -m raw

# 3. Verificar que aparezca
lpstat -p
```

> 🔧 **¿No usas USB?** Si la impresora está en red, el URI se ve como
> `ipp://192.168.1.x:631/ipp/print` o `socket://192.168.1.x:9100`.
> El asistente (`Instalar_Cola_Raw.command`) también te permite teclear el URI al elegir la opción "red".

---

## 4. Conectar el plugin con Ezy Ventas

1. Abre **Ezy Plugin de Impresión** (verás la banderita en la barra superior).
2. En tu sistema de **Ezy Ventas** ve a los ajustes del plugin de impresión.
3. Configura:
   - **Dirección del plugin:** `http://127.0.0.1:8000`
   - **Nombre de la impresora:** `EZY_Ticket`
4. Guarda y da clic en **Probar impresión**.

---

## 5. Prueba rápida manual (opcional)

Abre **Terminal** y ejecuta:

```bash
# Ver la versión del plugin
curl http://127.0.0.1:8000/version

# Ver las impresoras detectadas (debe aparecer EZY_Ticket)
curl http://127.0.0.1:8000/impresoras

# Imprimir un ticket de prueba
curl -X POST http://127.0.0.1:8000/imprimir \
  -H "Content-Type: application/json" \
  -d '{"nombreImpresora": "EZY_Ticket", "anchoImpresora": "80mm", "operaciones": [{"nombre": "EscribirTexto", "argumentos": ["\nPRUEBA DE TICKET MAC OS\n"]}, {"nombre": "Feed", "argumentos": [3]}]}'
```

Si todo funciona, la impresora debería imprimir el ticket de prueba.

---

## 6. Solución de problemas

| Problema | Solución |
|---|---|
| "No se puede abrir porque es de un desarrollador no verificado" | Clic derecho → **Abrir** → **Abrir** (paso 2). Es solo la primera vez. |
| `/impresoras` no muestra `EZY_Ticket` | Repite el paso 3 (cola raw) y verifica con `lpstat -p` en Terminal. |
| Imprime caracteres raros o símbolos | La impresora no está como cola raw → repite el paso 3. |
| No imprime pero el curl responde `ok` | Verifica que la impresora esté encendida y con papel. Revisa `Impresoras y escáneres`. |
| El cajón de dinero no abre | El cajón debe estar conectado **a la impresora** (puerto RJ11). Compatible con la mayoría de impresoras térmicas. |
| Ajustes del sistema no muestra la app | La app debe abrirse **una vez** (paso 2) antes de poder configurarla. |

---

## 7. Archivos de registro (para soporte)

Si algo falla, los registros del plugin están en:

```
/Usuarios/TU_USUARIO/Library/Application Support/EzyPlugin/impresion.log
```

Para abrir rápido en Finder: en Terminal escribe

```bash
open "$HOME/Library/Application Support/EzyPlugin"
```

📧 Envía el archivo `impresion.log` al soporte de Ezy Ventas si lo necesitan para ayudarte.