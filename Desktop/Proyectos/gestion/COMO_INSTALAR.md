# Cómo instalar Gestor de Turnos (macOS)

## Instalación fácil (recomendada)

1. Abre la carpeta del proyecto en el **Finder**.
2. Haz **doble clic** en `instalar.command`.
   - Si macOS avisa de que "no se puede abrir porque procede de un desarrollador
     no identificado", haz **clic derecho → Abrir** y confirma. Solo pasa la
     primera vez.
3. Espera a que termine (la primera vez tarda un poco: crea el entorno y compila).
4. Cuando veas **✅ ¡Listo!** ya tienes:
   - **Gestor Turnos** en la carpeta *Aplicaciones*.
   - Un **acceso directo en el Escritorio**.

Cada vez que quieras generar los turnos, abre la app desde el Escritorio.

---

## ¿Qué hace el instalador por dentro?

Es un script normal; puedes leerlo. En resumen:

```
python3 -m venv .venv           # entorno aislado
pip install -r requirements.txt # pandas, openpyxl, pyinstaller
pyinstaller "Gestor Turnos.spec"# genera "Gestor Turnos.app"
cp -R  ...  /Applications/       # instala la app
osascript ...                    # crea el alias en el Escritorio
```

## Compilar a mano (opcional)

Si prefieres hacerlo por tu cuenta desde la Terminal:

```bash
cd "ruta/al/proyecto"
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
pyinstaller --noconfirm "Gestor Turnos.spec"
open dist                       # ahí está "Gestor Turnos.app"
```

> **Nota:** PyInstaller genera un ejecutable para el sistema donde lo corres.
> El `.spec` incluido produce un **`.app` de macOS**; para Windows/Linux
> haría falta compilar en ese sistema.
