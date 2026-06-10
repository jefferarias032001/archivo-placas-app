@echo off
chcp 65001 >nul
REM ============================================================
REM  actualizar.bat — UN SOLO COMANDO para actualizar todo:
REM  1) Convierte el Excel de OneDrive en data.js
REM  2) Guarda el cambio en Git (commit)
REM  3) Lo sube a GitHub (push)
REM  Doble clic y listo.
REM ============================================================

cd /d "%~dp0"

echo.
echo [1/3] Convirtiendo el Excel de OneDrive a data.js ...
python convertir.py
if errorlevel 1 (
    echo.
    echo ❌ La conversion fallo. Revisa el mensaje de arriba.
    pause
    exit /b 1
)

echo.
echo [2/3] Guardando cambios en Git ...
git add data.js
git commit -m "Actualizacion de datos %date% %time%"
if errorlevel 1 (
    echo (No habia cambios nuevos para guardar, o revisa el mensaje de arriba)
)

echo.
echo [3/3] Subiendo a GitHub ...
git push

echo.
echo ✅ Listo. La app ya tiene los datos mas recientes.
pause
