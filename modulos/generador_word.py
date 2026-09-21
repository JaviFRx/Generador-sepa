"""
Módulo para generar consentimientos en formato PDF desde plantilla Word
"""

from docx import Document
from pathlib import Path
from typing import Dict
from datetime import datetime
import re
import os
from uuid import uuid4
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, DictionaryObject, ArrayObject, TextStringObject, NumberObject


class GeneradorConsentimientos:
    """Clase para generar documentos de consentimiento en PDF desde plantilla Word"""
    
    def __init__(self, ruta_plantilla: str, carpeta_salida: str):
        """
        Inicializa el generador de consentimientos
        
        Args:
            ruta_plantilla: Ruta a la plantilla Word
            carpeta_salida: Carpeta donde guardar los consentimientos generados
        """
        self.ruta_plantilla = Path(ruta_plantilla)
        self.carpeta_salida = Path(carpeta_salida)
        
        # Crear carpeta de salida si no existe
        self.carpeta_salida.mkdir(parents=True, exist_ok=True)
        
        # Verificar que Word esté disponible
        self._verificar_word_disponible()
    
    def _verificar_word_disponible(self):
        """Verifica que Word esté instalado y disponible"""
        import subprocess
        try:
            # Intentar crear una instancia de Word
            ps_test = '''
                try {
                    $word = New-Object -ComObject Word.Application
                    $word.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
                    Write-Output "OK"
                } catch {
                    Write-Error $_.Exception.Message
                    exit 1
                }
            '''
            result = subprocess.run(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_test],
                capture_output=True,
                timeout=10,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if result.returncode != 0:
                raise Exception("Word no está disponible o no tiene permisos")
                
        except subprocess.TimeoutExpired:
            raise Exception("Word no responde. Cierre Word si está abierto e intente de nuevo.")
        except Exception as e:
            print(f"⚠ Advertencia: No se pudo verificar Word - {str(e)}")
    
    def obtener_marcadores(self) -> list:
        """
        Obtiene todos los marcadores {{campo}} de la plantilla Word
        
        Returns:
            Lista de marcadores encontrados (sin las llaves)
        """
        doc = Document(self.ruta_plantilla)
        marcadores = set()
        
        # Buscar en párrafos
        for paragraph in doc.paragraphs:
            encontrados = re.findall(r'\{\{(\w+)\}\}', paragraph.text)
            marcadores.update(encontrados)
        
        # Buscar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        encontrados = re.findall(r'\{\{(\w+)\}\}', paragraph.text)
                        marcadores.update(encontrados)
        
        # Buscar en encabezados y pies de página
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                encontrados = re.findall(r'\{\{(\w+)\}\}', paragraph.text)
                marcadores.update(encontrados)
            for paragraph in section.footer.paragraphs:
                encontrados = re.findall(r'\{\{(\w+)\}\}', paragraph.text)
                marcadores.update(encontrados)
        
        return sorted(list(marcadores))
        
    def generar_consentimiento(self, datos: Dict) -> str:
        """
        Genera un documento de consentimiento en PDF con los datos proporcionados
        
        Los marcadores en la plantilla Word deben tener el formato:
        {{nombre_campo}} donde nombre_campo coincide con las claves del diccionario datos
        
        Ejemplos: {{nombre}}, {{apellido}}, {{email}}, {{dni}}, {{fecha}}
        
        Args:
            datos: Diccionario con los datos para reemplazar en la plantilla
            
        Returns:
            Ruta del archivo PDF generado
        """
        try:
            # Cargar la plantilla
            doc = Document(self.ruta_plantilla)
            
            # Reemplazar en párrafos
            for paragraph in doc.paragraphs:
                self._reemplazar_texto(paragraph, datos)
            
            # Reemplazar en tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            self._reemplazar_texto(paragraph, datos)
            
            # Reemplazar en encabezados y pies de página
            for section in doc.sections:
                # Encabezado
                for paragraph in section.header.paragraphs:
                    self._reemplazar_texto(paragraph, datos)
                # Pie de página
                for paragraph in section.footer.paragraphs:
                    self._reemplazar_texto(paragraph, datos)
            
            # Generar nombre de archivo (temporal Word)
            nombre_archivo_word = self._generar_nombre_archivo(datos)
            ruta_word = self.carpeta_salida / nombre_archivo_word
            
            # Guardar documento Word temporal
            doc.save(ruta_word)
            
            # Convertir a PDF
            nombre_archivo_pdf = nombre_archivo_word.replace('.docx', '.pdf')
            ruta_pdf = self.carpeta_salida / nombre_archivo_pdf
            
            # Usar PowerShell para conversión más rápida
            import subprocess
            import time
            
            # Script PowerShell inline con mejor manejo de errores
            ps_command = f'''
                $ErrorActionPreference = 'Stop'
                $word = $null
                try {{
                    # Crear nueva instancia de Word
                    $word = New-Object -ComObject Word.Application
                    $word.Visible = $false
                    $word.DisplayAlerts = 0
                    
                    # Abrir documento
                    $doc = $word.Documents.Open('{str(ruta_word.absolute())}', $false, $true)
                    
                    # Guardar como PDF
                    $doc.SaveAs([ref]'{str(ruta_pdf.absolute())}', [ref]17)
                    
                    # Cerrar documento
                    $doc.Close($false)
                    
                    Write-Output "PDF generado exitosamente"
                }} catch {{
                    Write-Error "Error: $($_.Exception.Message)"
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
                # Ejecutar PowerShell con timeout ampliado
                result = subprocess.run(
                    ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', ps_command],
                    capture_output=True,
                    timeout=45,  # Aumentado a 45 segundos
                    text=True,
                    creationflags=subprocess.CREATE_NO_WINDOW  # No mostrar ventana
                )
                
                # Esperar a que se escriba el archivo
                time.sleep(0.8)
                
                # Verificar si se generó el PDF
                if result.returncode != 0 or not ruta_pdf.exists() or ruta_pdf.stat().st_size == 0:
                    error_msg = result.stderr if result.stderr else "Desconocido"
                    raise Exception(f"PDF no generado. Error: {error_msg[:300]}")
                
                print(f"  ✓ PDF generado correctamente")
                    
            except subprocess.TimeoutExpired:
                # Si hay timeout, intentar matar Word y verificar si el PDF se generó
                print(f"  ⚠ Timeout detectado, intentando recuperar...")
                try:
                    subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                                 capture_output=True, timeout=3)
                    time.sleep(2)
                    if ruta_pdf.exists():
                        print(f"  ✓ PDF generado (Word cerrado forzosamente)")
                    else:
                        raise Exception("Timeout: Word tardó más de 45 segundos. Por favor:\n"
                                      "1. Cierre Word manualmente si está abierto\n"
                                      "2. Verifique que Word esté instalado correctamente\n"
                                      "3. Intente generar menos PDFs a la vez")
                except:
                    raise Exception("Timeout al convertir a PDF.")
            except Exception as e:
                if "PDF no generado" in str(e):
                    raise
                raise Exception(f"Error al convertir Word a PDF: {str(e)}")
            finally:
                # Limpiar siempre - CRÍTICO
                try:
                    subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                                 capture_output=True, timeout=2)
                    time.sleep(0.3)
                except:
                    pass
                
                # Eliminar archivo Word temporal siempre
                try:
                    if ruta_word.exists():
                        os.remove(ruta_word)
                except:
                    pass
            
            # Añadir campos editables al PDF
            self._anadir_campos_editables(ruta_pdf)
            
            return str(ruta_pdf)
            
        except Exception as e:
            raise Exception(f"Error al generar consentimiento: {str(e)}")
    
    def _reemplazar_texto(self, paragraph, datos: Dict):
        """
        Reemplaza los marcadores en un párrafo con los datos proporcionados
        
        Args:
            paragraph: Párrafo del documento
            datos: Diccionario con los datos
        """
        # Obtener todo el texto del párrafo
        texto_completo = paragraph.text
        
        # Buscar todos los marcadores {{campo}}
        marcadores = re.findall(r'\{\{(\w+)\}\}', texto_completo)
        
        if not marcadores:
            return
        
        # Crear el nuevo texto reemplazando marcadores
        texto_nuevo = texto_completo
        for marcador in marcadores:
            # Buscar el valor en los datos (insensible a mayúsculas)
            valor = None
            for clave, val in datos.items():
                if clave.lower() == marcador.lower():
                    valor = str(val) if val is not None else ""
                    break
            
            if valor is not None:
                texto_nuevo = texto_nuevo.replace(f"{{{{{marcador}}}}}", valor)
            else:
                # Si no se encuentra el valor, dejar vacío
                texto_nuevo = texto_nuevo.replace(f"{{{{{marcador}}}}}", "")
        
        # Reemplazar todo el texto del párrafo
        if texto_nuevo != texto_completo:
            # Borrar todos los runs existentes
            for run in paragraph.runs:
                run.text = ""
            
            # Si no hay runs, crear uno
            if not paragraph.runs:
                paragraph.add_run(texto_nuevo)
            else:
                # Poner todo el texto en el primer run
                paragraph.runs[0].text = texto_nuevo
    
    def _generar_nombre_archivo(self, datos: Dict) -> str:
        """
        Genera un nombre de archivo basado en los datos
        
        Args:
            datos: Diccionario con los datos
            
        Returns:
            Nombre del archivo
        """
        # Intentar usar nombre y apellido, o DNI, o un identificador genérico
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Buscar campos comunes para el nombre (deudor)
        nombre = None
        apellido = None
        dni = None
        nombre_alumno = None
        
        for clave, valor in datos.items():
            clave_lower = clave.lower()
            if 'nombre_alumno' in clave_lower:
                nombre_alumno = valor
            elif 'nombre' in clave_lower and 'apellido' not in clave_lower:
                nombre = valor
            elif 'apellido' in clave_lower and 'alumno' not in clave_lower:
                apellido = valor
            elif 'dni' in clave_lower or 'documento' in clave_lower:
                dni = valor
        
        # Construir nombre de archivo
        partes = []
        if apellido:
            partes.append(self._limpiar_nombre(apellido))
        if nombre:
            partes.append(self._limpiar_nombre(nombre))
        if dni:
            partes.append(self._limpiar_nombre(dni))
        
        # Añadir nombre del alumno para diferenciar entre hermanos
        if nombre_alumno:
            partes.append(self._limpiar_nombre(nombre_alumno))
        
        if partes:
            nombre_base = "_".join(partes)
        else:
            nombre_base = f"consentimiento_{timestamp}"
        
        # Evitar sobrescribir PDFs de homónimos, hermanos o generaciones anteriores.
        return f"{nombre_base[:120]}_{uuid4().hex}.docx"
    
    def _limpiar_nombre(self, texto: str) -> str:
        """
        Limpia un texto para usarlo en nombre de archivo
        
        Args:
            texto: Texto a limpiar
            
        Returns:
            Texto limpio
        """
        # Remover caracteres no válidos para nombres de archivo
        texto = re.sub(r'[<>:"/\\|?*]', '', texto)
        # Reemplazar espacios por guiones bajos
        texto = texto.replace(' ', '_')
        # Limitar longitud
        texto = texto[:50]
        return texto
    
    def _anadir_campos_editables(self, ruta_pdf: Path):
        """
        Añade campos de formulario editables al PDF
        
        Args:
            ruta_pdf: Ruta del archivo PDF al que añadir los campos
        """
        try:
            # Leer el PDF existente
            reader = PdfReader(str(ruta_pdf))
            writer = PdfWriter()
            
            # Copiar todas las páginas
            for page in reader.pages:
                writer.add_page(page)
            
            # Obtener la última página (donde están fecha, localidad y firma)
            ultima_pagina = len(reader.pages) - 1
            page_height = float(reader.pages[ultima_pagina].mediabox.height)
            page_width = float(reader.pages[ultima_pagina].mediabox.width)
            
            # Posiciones aproximadas para los campos (ajustar según tu plantilla)
            # Estos valores son aproximados y pueden necesitar ajuste
            campos = [
                {
                    'nombre': 'fecha',
                    'x': 150,
                    'y': page_height - 680,  # Ajustar según posición en tu plantilla
                    'width': 120,
                    'height': 20
                },
                {
                    'nombre': 'localidad',
                    'x': 350,
                    'y': page_height - 680,  # Misma línea que fecha
                    'width': 200,
                    'height': 20
                },
                {
                    'nombre': 'firma',
                    'x': 150,
                    'y': page_height - 720,  # Debajo de fecha/localidad
                    'width': 400,
                    'height': 60
                }
            ]
            
            # Añadir campos al PDF
            writer.add_form_field(
                self._crear_campo_texto('fecha', 150, page_height - 680, 120, 20, ultima_pagina)
            )
            writer.add_form_field(
                self._crear_campo_texto('localidad', 350, page_height - 680, 200, 20, ultima_pagina)
            )
            writer.add_form_field(
                self._crear_campo_texto('firma', 150, page_height - 720, 400, 60, ultima_pagina)
            )
            
            # Guardar el PDF con los campos editables
            with open(str(ruta_pdf), 'wb') as output_file:
                writer.write(output_file)
        
        except Exception as e:
            # Si falla la adición de campos, el PDF ya está generado, solo avisar
            print(f"Advertencia: No se pudieron añadir campos editables: {str(e)}")
    
    def _crear_campo_texto(self, nombre, x, y, width, height, pagina):
        """
        Crea un campo de texto editable para el formulario PDF
        
        Args:
            nombre: Nombre del campo
            x, y: Posición en la página
            width, height: Dimensiones del campo
            pagina: Número de página donde colocar el campo
            
        Returns:
            Objeto de campo de formulario PDF
        """
        from pypdf.generic import NameObject, DictionaryObject, ArrayObject, TextStringObject, NumberObject
        
        # Crear anotación de widget
        widget = DictionaryObject()
        widget.update({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/Rect"): ArrayObject([
                NumberObject(x),
                NumberObject(y),
                NumberObject(x + width),
                NumberObject(y + height)
            ]),
            NameObject("/FT"): NameObject("/Tx"),  # Tipo: Text
            NameObject("/T"): TextStringObject(nombre),  # Nombre del campo
            NameObject("/V"): TextStringObject(""),  # Valor inicial vacío
            NameObject("/P"): NumberObject(pagina),  # Página
            NameObject("/F"): NumberObject(4),  # Flags: imprimible
        })
        
        return widget
