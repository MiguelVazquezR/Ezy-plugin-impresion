# -*- mode: python ; coding: utf-8 -*-
"""
ezy_plugin_macos.spec — Configuración de PyInstaller para macOS.

Genera un binario UNIVERSAL2 (Apple Silicon M1/M2/M3 + Intel) empaquetado
como aplicación .app. El workflow de GitHub Actions lo comprime en un .zip
listo para distribuir por WhatsApp/e-mail.

Uso (desde una Mac o runner macOS):
    pyinstaller ezy_plugin_macos.spec --noconfirm
"""

import os
import escpos

# Arquitectura de compilación. Por defecto universal2 (Intel + Apple Silicon).
# Se puede sobreescribir con EZYPLUGIN_ARCH=arm64 o EZYPLUGIN_ARCH=x86_64.
TARGET_ARCH = os.environ.get('EZYPLUGIN_ARCH', 'universal2')

# Ícono de la app (.icns). El workflow lo genera desde icon.png con sips +
# iconutil; si no existe, la app usa el ícono genérico de macOS.
ICON_PATH = os.path.join(SPECPATH, 'icon.icns')
APP_ICON = ICON_PATH if os.path.exists(ICON_PATH) else None

# Ruta portable del capabilities.json de python-escpos (varía entre sistemas).
capabilities_path = os.path.join(os.path.dirname(escpos.__file__),
                                 'capabilities.json')

a = Analysis(
    ['mi_servidor_impresion.py'],
    pathex=[],
    binaries=[],
    datas=[(capabilities_path, 'escpos'), ('icon.png', '.')],
    hiddenimports=[
        # pystray en macOS usa PyObjC (instalado vía pyobjc-framework-Cocoa).
        # Se declaran aquí explícitamente para que PyInstaller los incluya.
        'AppKit',
        'Foundation',
        'objc',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ezy_plugin_macos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX no aplica a binarios de macOS (evita firmas rotas)
    console=False,      # App sin ventana de terminal
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=TARGET_ARCH,    # universal2 (por defecto): Intel + Apple Silicon
    codesign_identity=None,     # Sin Developer ID (se abre con clic derecho → Abrir)
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    name='ezy_plugin_macos.app',
    icon=APP_ICON,
    bundle_identifier='com.ezyventas.ezyplugin',
    info_plist={
        'CFBundleName': 'Ezy Plugin de Impresión',
        'CFBundleDisplayName': 'Ezy Plugin de Impresión',
        'CFBundleShortVersionString': '1.2.0',
        'CFBundleVersion': '1.2.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '11.0',
    },
)