@echo off
REM ============================================================
REM  actualizar.bat - UN SOLO COMANDO para actualizar todo:
REM  1. Convierte el Excel de placas (OneDrive) en data.js
REM  2. Convierte la carpeta de viajes (OneDrive) en data_viajes.js
REM  3. Sube todos los cambios a GitHub
REM ============================================================

cd /d "%~dp0"

echo.
echo [1/4] Convirtiendo el archivo de placas a data.js ...
python convertir.py
if errorlevel 1 (
    echo.
    echo [ERROR] La conversion de placas fallo. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo [2/4] Convirtiendo la carpeta de viajes a data_viajes.js ...
python convertir_viajes.py
if errorlevel 1 (
    echo.
    echo [ERROR] La conversion de viajes fallo. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo [3/4] Guardando cambios en Git ...
git add .
git commit -m "Actualizacion de datos %date% %time%"

echo.
echo [4/4] Subiendo a GitHub ...
git pull
git push

echo.
echo [OK] Listo. La app ya tiene los datos mas recientes.
pause
