#!/bin/bash
#
# Gestor de Turnos — instalador para macOS
# --------------------------------------------------------------------
# Doble clic en este archivo (o ejecútalo desde Terminal) para:
#   1. Crear un entorno virtual e instalar las dependencias.
#   2. Compilar la app con PyInstaller  ->  "Gestor Turnos.app".
#   3. Copiarla a /Applications y dejar un acceso directo en el Escritorio.
#
# No necesitas saber nada de terminal: solo doble clic y esperar.
# --------------------------------------------------------------------
set -e

# Situarse en la carpeta del script (funciona con doble clic)
cd "$(dirname "$0")"

echo "======================================"
echo "  Gestor de Turnos — instalación"
echo "======================================"
echo

# 1. Comprobar Python 3 ------------------------------------------------
if ! command -v python3 >/dev/null 2>&1; then
    echo "❌ No se encuentra Python 3."
    echo "   Instálalo desde https://www.python.org/downloads/ y vuelve a intentarlo."
    read -n 1 -s -r -p "Pulsa una tecla para cerrar…"
    exit 1
fi
echo "✅ Python 3 detectado: $(python3 --version)"

# 2. Entorno virtual + dependencias -----------------------------------
echo
echo "▶ Preparando dependencias (esto puede tardar la primera vez)…"
python3 -m venv .venv
source .venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# 3. Compilar ----------------------------------------------------------
echo
echo "▶ Compilando la aplicación…"
rm -rf build dist
pyinstaller --noconfirm "Gestor Turnos.spec"

APP="dist/Gestor Turnos.app"
if [ ! -d "$APP" ]; then
    echo "❌ No se generó la app. Revisa los mensajes anteriores."
    read -n 1 -s -r -p "Pulsa una tecla para cerrar…"
    exit 1
fi

# 4. Instalar en /Applications ----------------------------------------
echo
echo "▶ Instalando en Aplicaciones…"
rm -rf "/Applications/Gestor Turnos.app"
cp -R "$APP" "/Applications/"

# 5. Acceso directo (alias) en el Escritorio --------------------------
echo "▶ Creando acceso directo en el Escritorio…"
DESKTOP="$HOME/Desktop"
osascript >/dev/null <<'APPLESCRIPT'
tell application "Finder"
    set appFile to POSIX file "/Applications/Gestor Turnos.app" as alias
    set desktopFolder to path to desktop folder
    try
        delete file "Gestor Turnos alias" of desktopFolder
    end try
    make new alias file at desktopFolder to appFile
    set name of result to "Gestor Turnos"
end tell
APPLESCRIPT

echo
echo "======================================"
echo "  ✅ ¡Listo!"
echo "  • App instalada en Aplicaciones."
echo "  • Acceso directo en tu Escritorio."
echo "======================================"
echo
read -n 1 -s -r -p "Pulsa una tecla para cerrar…"
