@echo off
echo ========================================
echo    LIMPIEZA DE PROCESOS
echo ========================================
echo.
echo Cerrando Word...
taskkill /F /IM WINWORD.EXE 2>nul
if %errorlevel% equ 0 (
    echo   OK - Word cerrado
) else (
    echo   - No habia procesos de Word
)
echo.
echo Cerrando Python bloqueados...
tasklist | find /i "python.exe" >nul
if %errorlevel% equ 0 (
    echo   Se detectaron procesos Python:
    tasklist | find /i "python.exe"
    echo.
    echo   Si la aplicacion esta bloqueada, presione Ctrl+C en la ventana
) else (
    echo   - No hay procesos Python
)
echo.
echo ========================================
echo    Listo
echo ========================================
pause
