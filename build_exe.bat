@echo off
cd /d "%~dp0"

echo Instalando/actualizando PyInstaller y dependencias...
pip install -r requirements.txt pyinstaller

echo.
echo Generando ejecutable standalone...
python -m PyInstaller --onefile --windowed --name "AutomatizadorInventariosSIESAM" main.py

echo.
echo Listo. El ejecutable quedo en: dist\AutomatizadorInventariosSIESAM.exe
pause
