@echo off
title Academia Horizonte - Sistema de Certificados
cd /d "%~dp0"

rem Ruta de Python (no esta en el PATH de Windows por defecto)
set "PATH=C:\Users\Natasha\AppData\Local\Programs\Python\Python312;C:\Users\Natasha\AppData\Local\Programs\Python\Python312\Scripts;%PATH%"

echo ============================================
echo  Academia Horizonte - Sistema de Certificados
echo  Abre tu navegador en: http://127.0.0.1:5000
echo  Para DETENER la app: presiona Ctrl+C aqui
echo ============================================
echo.

python app.py

echo.
echo La aplicacion se cerro. Presiona una tecla para salir.
pause > nul
