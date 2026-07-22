@echo off
title Lanzador Aplicacion SPAE Standalone
echo ===================================================
echo   INICIANDO SERVIDOR Y APLICACION SPAE STANDALONE
echo ===================================================
echo.
cd /d C:\Users\cagch\Desktop\SPAE_Aislada

echo Abriendo navegador en http://localhost:8505 ...
start "" http://localhost:8505

echo Ejecutando aplicacion Streamlit en Puerto 8505...
python -m streamlit run app.py --server.port 8505

pause
