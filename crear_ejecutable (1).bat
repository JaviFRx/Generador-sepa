@echo off
echo ========================================
echo   Generando Ejecutable .EXE
echo   Generador de Consentimientos SEPAS
echo ========================================
echo.

REM Activar entorno virtual
call .venv\Scripts\activate.bat

echo Compilando aplicacion...
echo Esto puede tardar 1-2 minutos...
echo.

pyinstaller --onefile --windowed --name="Generador_Consentimientos_SEPAS" --add-data "modulos;modulos" app_consentimientos.py

echo.
echo ========================================
echo   Compilacion completada!
echo ========================================
echo.
echo El archivo .EXE se encuentra en:
echo   dist\Generador_Consentimientos_SEPAS.exe
echo.
echo Puedes copiar este archivo a cualquier PC con Windows
echo y ejecutarlo sin necesidad de instalar Python.
echo.
echo IMPORTANTE: Cuando uses el .exe, asegurate de tener
echo la carpeta 'modulos' en el mismo directorio.
echo.
pause
