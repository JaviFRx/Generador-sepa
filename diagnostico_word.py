"""
Script de diagnóstico para verificar la instalación de Word
"""

import subprocess
import sys

print("=" * 80)
print("DIAGNÓSTICO DE MICROSOFT WORD PARA CONVERSIÓN PDF")
print("=" * 80)

# 1. Verificar que Word está instalado
print("\n1. Verificando instalación de Word...")
ps_check_word = '''
try {
    $word = New-Object -ComObject Word.Application
    Write-Output "Versión de Word: $($word.Version)"
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
    Write-Error "Error: $_"
    exit 1
}
'''

try:
    result = subprocess.run(
        ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_check_word],
        capture_output=True,
        timeout=15,
        text=True
    )
    
    if result.returncode == 0:
        print(f"✓ Word instalado: {result.stdout.strip()}")
    else:
        print(f"✗ Error al acceder a Word:")
        print(f"  {result.stderr}")
        sys.exit(1)
except subprocess.TimeoutExpired:
    print("✗ Timeout: Word no responde en 15 segundos")
    sys.exit(1)
except Exception as e:
    print(f"✗ Error: {e}")
    sys.exit(1)

# 2. Verificar procesos de Word en ejecución
print("\n2. Verificando procesos de Word...")
ps_check_process = "Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object Id, ProcessName"

result = subprocess.run(
    ['powershell', '-Command', ps_check_process],
    capture_output=True,
    text=True
)

if result.stdout.strip():
    print("⚠ ADVERTENCIA: Word está en ejecución:")
    print(result.stdout)
    print("  Recomendación: Cierre Word antes de generar PDFs")
else:
    print("✓ No hay procesos de Word en ejecución")

# 3. Test de conversión simple
print("\n3. Probando conversión Word → PDF...")
print("  (Esto puede tardar unos segundos...)")

from pathlib import Path
from docx import Document

# Crear documento de prueba
test_dir = Path("test_conversion")
test_dir.mkdir(exist_ok=True)

doc = Document()
doc.add_paragraph("Documento de prueba para conversión PDF")
test_word = test_dir / "test.docx"
test_pdf = test_dir / "test.pdf"

doc.save(test_word)

ps_convert = f'''
$ErrorActionPreference = 'Stop'
$word = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    
    $doc = $word.Documents.Open('{str(test_word.absolute())}', $false, $true)
    $doc.SaveAs([ref]'{str(test_pdf.absolute())}', [ref]17)
    $doc.Close($false)
    
    Write-Output "Conversión exitosa"
}} catch {{
    Write-Error "Error en conversión: $($_.Exception.Message)"
    exit 1
}} finally {{
    if ($word -ne $null) {{
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    }}
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}}
'''

try:
    result = subprocess.run(
        ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_convert],
        capture_output=True,
        timeout=30,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW
    )
    
    import time
    time.sleep(1)
    
    if test_pdf.exists():
        print("✓ Conversión de prueba exitosa")
        print(f"  Archivo generado: {test_pdf}")
        
        # Limpiar archivos de prueba
        import shutil
        shutil.rmtree(test_dir)
        print("✓ Archivos de prueba eliminados")
    else:
        print("✗ El PDF no se generó")
        print(f"  Salida: {result.stdout}")
        print(f"  Error: {result.stderr}")
except subprocess.TimeoutExpired:
    print("✗ Timeout en conversión de prueba (>30 segundos)")
    print("  Esto indica que Word está muy lento o bloqueado")
except Exception as e:
    print(f"✗ Error en prueba: {e}")

print("\n" + "=" * 80)
print("DIAGNÓSTICO COMPLETADO")
print("=" * 80)
print("\nRecomendaciones:")
print("1. Si hay procesos de Word, ciérrelos antes de generar PDFs")
print("2. Si la conversión de prueba falla, reinicie el ordenador")
print("3. Verifique que Word está activado y actualizado")
print("4. Intente generar PDFs en lotes pequeños (5-10 a la vez)")
