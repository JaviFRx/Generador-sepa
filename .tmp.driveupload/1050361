"""
Módulo para generar consentimientos en formato PDF desde plantilla Word
"""

from docx import Document
from docx2pdf import convert
from pathlib import Path
from typing import Dict
from datetime import datetime
import re
import os


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
            
            try:
                convert(str(ruta_word), str(ruta_pdf))
            except Exception as e:
                # Si falla la conversión, dar más información
                raise Exception(f"Error al convertir Word a PDF: {str(e)}. "
                              f"Asegúrese de que Microsoft Word esté instalado y cerrado. "
                              f"El archivo Word temporal está en: {ruta_word}")
            
            # Eliminar archivo Word temporal solo si la conversión fue exitosa
            if ruta_pdf.exists():
                os.remove(ruta_word)
            
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
        
        # Buscar campos comunes para el nombre
        nombre = None
        apellido = None
        dni = None
        
        for clave, valor in datos.items():
            clave_lower = clave.lower()
            if 'nombre' in clave_lower and not 'apellido' in clave_lower:
                nombre = valor
            elif 'apellido' in clave_lower:
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
        
        if partes:
            nombre_base = "_".join(partes)
        else:
            nombre_base = f"consentimiento_{timestamp}"
        
        return f"{nombre_base}.docx"
    
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
