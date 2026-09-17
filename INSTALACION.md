# Guía de Instalación — Automatizador de Inventarios SI ESAM

Instrucciones paso a paso para instalar y correr la aplicación en tu computadora (Windows), partiendo de cero. No necesitas saber programar.

---

## Paso 1 — Descargar el proyecto desde GitHub

1. Ve a: `https://github.com/Neycer-Nay-99/AUTOMATIZADOR-INVENTARIOS-SiESAM`
2. Clic en el botón verde **"Code"** → **"Download ZIP"**.
3. Descomprime el ZIP en una carpeta fácil de encontrar, por ejemplo `C:\Automatizador Inventarios`.

---

## Paso 2 — Instalar Python

Si ya tienes Python 3.10 o superior instalado, salta al **Paso 3**. Para verificar, abre **PowerShell** (busca "PowerShell" en el menú de inicio) y escribe:
```powershell
python --version
```
Si te dice algo como `Python 3.11.x`, ya lo tienes. Si te da error o no lo reconoce:

1. Ve a `https://www.python.org/downloads/`
2. Descarga la versión más reciente para Windows.
3. Al ejecutar el instalador, **muy importante**: marca la casilla **"Add python.exe to PATH"** antes de darle a "Install Now".
4. Termina la instalación y cierra/vuelve a abrir PowerShell.
5. Verifica de nuevo con `python --version`.

---

## Paso 3 — Instalar las dependencias del proyecto

1. Abre **PowerShell**.
2. Entra a la carpeta donde descomprimiste el proyecto, por ejemplo:
   ```powershell
   cd "C:\Automatizador Inventarios"
   ```
3. Crea un entorno virtual (esto mantiene las librerías del proyecto aisladas del resto de tu sistema):
   ```powershell
   python -m venv venv
   ```
4. Activa el entorno virtual:
   ```powershell
   .\venv\Scripts\activate
   ```
   Si ves un error de permisos ("no se puede cargar el archivo... porque la ejecución de scripts está deshabilitada"), ejecuta esto una sola vez y vuelve a intentar:
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
   ```
5. Instala las librerías necesarias:
   ```powershell
   pip install -r requirements.txt
   ```
   Esto instala `pandas` y `openpyxl`, las únicas librerías necesarias.

---

## Paso 4 — Ejecutar la aplicación

Ya no necesitas PowerShell para uso diario. Simplemente:

- **Doble clic en `lanzar.bat`** (dentro de la carpeta del proyecto).

Esto abre la ventana de la aplicación. Para aprender a usarla, revisa **`MANUAL_USUARIO.md`**.

> Nota: `lanzar.bat` detecta automáticamente el entorno virtual (`venv`) que creaste en el Paso 3 y lo usa; si por algún motivo no existe, intenta usar el Python global de tu sistema.

---

## Opción alternativa — Generar un ejecutable (.exe) para usar en otras PC sin instalar Python

Si necesitás correr la app en otra computadora sin repetir los pasos 1-3 (instalar Python, crear entorno virtual, etc.), podés generar un `.exe` standalone que funciona con doble clic en cualquier PC con Windows, sin instalar nada.

**En la PC donde ya tenés el proyecto instalado (pasos 1-3 ya hechos):**

1. Doble clic en **`build_exe.bat`**.
2. Esperá unos minutos — instala `PyInstaller` y genera el ejecutable.
3. El resultado queda en `dist\AutomatizadorInventariosSIESAM.exe`.

**Para usarlo en otra PC:** copiá solo ese archivo `.exe` (no hace falta copiar el resto del proyecto) y hacé doble clic — se abre la misma aplicación, sin necesidad de instalar Python ni ninguna librería.

> Notas:
> - El `.exe` pesa varios MB (~90 MB) porque incluye Python y todas las librerías (pandas, openpyxl) empaquetadas adentro — es normal.
> - Cada vez que se modifique `main.py`, hay que volver a correr `build_exe.bat` para regenerar el `.exe` con los cambios.
> - El primer arranque del `.exe` puede tardar unos segundos más que `lanzar.bat` (tiene que descomprimirse en memoria); los siguientes arranques son más rápidos.

---

## Problemas comunes

| Problema | Solución |
|----------|----------|
| `'python' no se reconoce como comando` | Python no quedó en el PATH — reinstala marcando "Add python.exe to PATH" (Paso 2) |
| `ModuleNotFoundError: No module named 'pandas'` | No se activó el entorno virtual antes de instalar, o faltó ejecutar `pip install -r requirements.txt` (Paso 3) |
| La ventana no abre al hacer doble clic en `lanzar.bat` | Abre PowerShell, entra a la carpeta del proyecto y ejecuta `python main.py` directamente para ver el mensaje de error exacto |
| No se puede activar el entorno virtual (error de "scripts deshabilitados") | Ejecuta `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` en PowerShell (Paso 3, punto 4) |

---

Para aprender a usar la aplicación una vez instalada, ver **`MANUAL_USUARIO.md`**.
Para detalles técnicos del código (personalización de categorías/marcas, estructura de columnas), ver **`README.md`**.
