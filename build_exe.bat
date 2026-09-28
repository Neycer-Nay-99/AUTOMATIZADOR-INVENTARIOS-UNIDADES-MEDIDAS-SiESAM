@echo off
cd /d "%~dp0"

echo Instalando/actualizando PyInstaller y dependencias...
pip install -r requirements.txt pyinstaller

echo.
echo Generando ejecutable standalone...
python -m PyInstaller --onefile --windowed --icon "siesam.ico" --add-data "siesam.ico;." --name "SIESAM_Inventarios" main.py

echo.
echo Listo. El ejecutable quedo en: dist\SIESAM_Inventarios.exe
pause
