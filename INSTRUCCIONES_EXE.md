# Ejecutable de Windows

## Ubicación y uso

La entrega se instala directamente en `E:/Sepas/dist/Generador_Consentimientos_SEPAS.exe`.
No se utiliza `dist/actualizado`. Sigue la [guía rápida](GUIA_RAPIDA.md).

El archivo incluye Python y las dependencias. Microsoft Word debe estar instalado
para generar PDFs. La cuenta Gmail se configura en cada equipo.
La configuración está en `.sepas_config.json`, dentro del perfil de Windows.
Guarda los PDFs y su historial fuera de `dist`, porque se reemplaza al actualizar.

## Compilar

Desde la raíz del proyecto, en PowerShell:

```powershell
.venv/Scripts/python.exe -m pip install -r requirements.txt
.venv/Scripts/python.exe -m pip install pyinstaller
.venv/Scripts/python.exe -m PyInstaller --noconfirm --workpath build/progreso --distpath build/entrega_progreso Generador_Consentimientos_SEPAS.spec
```

La entrada es `app_consentimientos.py`. El `.spec` principal incorpora los módulos
y la plantilla de `Docs`. Los archivos con sufijo `(1)` son copias históricas,
no las entradas de la compilación actual.

En este equipo se usa una copia local de Tcl/Tk para compilar y ejecutar las
pruebas gráficas. Si esas carpetas están presentes, configura antes:

```powershell
$env:TCL_LIBRARY = 'E:/Sepas/build/tcl_runtime/tcl8.6'
$env:TK_LIBRARY = 'E:/Sepas/build/tcl_runtime/tk8.6'
$env:PYINSTALLER_CONFIG_DIR = 'E:/Sepas/build/pyinstaller_config'
```

Son rutas específicas de este entorno; no hacen falta cuando Python dispone de
Tcl/Tk accesible normalmente. Las entregas documentadas se compilaron con
Python 3.10.10 y PyInstaller 6.18.0.

## Reemplazar la entrega

1. Compila en `build/entrega_progreso` y comprueba que finaliza sin errores.
2. Ejecuta las [pruebas](DOCUMENTACION_TECNICA.md#pruebas) y verifica que el archivo
   incluye los módulos actualizados.
3. Cierra la aplicación: Windows impide reemplazar un ejecutable en uso.
4. Comprueba que el destino resuelto es exactamente `E:/Sepas/dist`, sin enlaces
   o uniones a otras carpetas y con solo archivos de entrega.
5. Elimina esa carpeta, créala de nuevo y copia dentro el ejecutable compilado.
6. Compara las huellas SHA-256 del archivo compilado y la copia final.

La eliminación se limita a `dist`. Conserva las carpetas de salida, historiales,
plantillas elegidas por el usuario y configuración del perfil.

`dist/` y `build/` están excluidos de Git. El push publica código y documentación;
el ejecutable se distribuye por separado. Compilar no envía correos.
