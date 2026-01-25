@echo off
echo ========================================
echo    PREPARANDO GENERACION DE PDFs
echo ========================================
echo.
echo Cerrando procesos de Word...
taskkill /F /IM WINWORD.EXE 2>nul
timeout /t 2 /nobreak >nul
echo.
echo Iniciando aplicacion...
python app_consentimientos.py
pause
