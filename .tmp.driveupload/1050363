"""
Módulo para preparar emails con consentimientos adjuntos
"""

from pathlib import Path
from typing import List, Dict
import openpyxl


class PreparadorEmails:
    """Clase para preparar emails con consentimientos"""
    
    def __init__(self, carpeta_consentimientos: str):
        """
        Inicializa el preparador de emails
        
        Args:
            carpeta_consentimientos: Carpeta donde están los consentimientos generados
        """
        self.carpeta_consentimientos = Path(carpeta_consentimientos)
        
    def crear_borradores_email(self, ruta_excel: str) -> List[Dict]:
        """
        Crea borradores de email basados en el Excel y los consentimientos generados
        
        Args:
            ruta_excel: Ruta al archivo Excel con los datos
            
        Returns:
            Lista de diccionarios con información de cada email
        """
        from modulos.lector_excel import LectorExcel
        
        # Leer datos del Excel
        lector = LectorExcel(ruta_excel)
        datos = lector.leer_datos()
        
        # Obtener lista de archivos PDF generados
        archivos_generados = list(self.carpeta_consentimientos.glob("*.pdf"))
        
        # Crear borradores
        borradores = []
        archivo_salida = self.carpeta_consentimientos / "emails_para_enviar.txt"
        
        with open(archivo_salida, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("BORRADORES DE EMAILS PARA ENVÍO DE CONSENTIMIENTOS (PDF)\n")
            f.write(f"Generado: {Path(ruta_excel).name}\n")
            f.write("=" * 80 + "\n\n")
            
            for i, registro in enumerate(datos, 1):
                # Buscar email en el registro
                email = self._encontrar_email(registro)
                nombre = self._encontrar_nombre(registro)
                
                # Buscar archivo de consentimiento correspondiente
                archivo_consentimiento = self._encontrar_consentimiento(registro, archivos_generados)
                
                # Crear borrador
                borrador = {
                    'numero': i,
                    'email': email,
                    'nombre': nombre,
                    'archivo': archivo_consentimiento,
                    'asunto': f"Consentimiento - {nombre}",
                    'cuerpo': self._generar_cuerpo_email(nombre)
                }
                
                borradores.append(borrador)
                
                # Escribir en archivo
                f.write(f"EMAIL #{i}\n")
                f.write("-" * 80 + "\n")
                f.write(f"Para: {email}\n")
                f.write(f"Asunto: {borrador['asunto']}\n")
                f.write(f"Adjunto: {Path(archivo_consentimiento).name if archivo_consentimiento else 'NO ENCONTRADO'}\n")
                f.write(f"\nCuerpo del mensaje:\n")
                f.write(borrador['cuerpo'])
                f.write("\n\n" + "=" * 80 + "\n\n")
        
        # Crear archivo CSV para importar a cliente de email (opcional)
        self._crear_csv_emails(borradores)
        
        return borradores
    
    def _encontrar_email(self, registro: Dict) -> str:
        """Encuentra el email en el registro"""
        for clave, valor in registro.items():
            if 'email' in clave.lower() or 'correo' in clave.lower() or 'mail' in clave.lower():
                return valor
        return "email@no-encontrado.com"
    
    def _encontrar_nombre(self, registro: Dict) -> str:
        """Encuentra el nombre completo en el registro"""
        nombre = ""
        apellido = ""
        
        for clave, valor in registro.items():
            clave_lower = clave.lower()
            if 'nombre' in clave_lower and 'apellido' not in clave_lower:
                nombre = valor
            elif 'apellido' in clave_lower:
                apellido = valor
        
        if apellido and nombre:
            return f"{nombre} {apellido}"
        elif nombre:
            return nombre
        elif apellido:
            return apellido
        return "Estimado/a"
    
    def _encontrar_consentimiento(self, registro: Dict, archivos: List[Path]) -> str:
        """
        Encuentra el archivo de consentimiento correspondiente a un registro
        
        Args:
            registro: Registro con los datos
            archivos: Lista de archivos generados
            
        Returns:
            Ruta del archivo o None
        """
        # Buscar DNI o identificador
        dni = None
        nombre = None
        apellido = None
        
        for clave, valor in registro.items():
            clave_lower = clave.lower()
            if 'dni' in clave_lower or 'documento' in clave_lower:
                dni = valor
            elif 'nombre' in clave_lower and 'apellido' not in clave_lower:
                nombre = valor
            elif 'apellido' in clave_lower:
                apellido = valor
        
        # Buscar archivo que contenga el DNI, apellido o nombre
        for archivo in archivos:
            nombre_archivo = archivo.stem.lower()
            
            if dni and dni.lower() in nombre_archivo:
                return str(archivo)
            if apellido and apellido.lower() in nombre_archivo:
                return str(archivo)
            if nombre and nombre.lower() in nombre_archivo:
                return str(archivo)
        
        return None
    
    def _generar_cuerpo_email(self, nombre: str) -> str:
        """
        Genera el cuerpo del email
        
        Args:
            nombre: Nombre del destinatario
            
        Returns:
            Cuerpo del email
        """
        return f"""Estimado/a {nombre},

Adjunto encontrará el documento de consentimiento para su revisión y firma.

Por favor, revise el documento cuidadosamente. Si está de acuerdo con los términos, 
le solicitamos que lo firme y nos lo envíe de vuelta.

Si tiene alguna pregunta o inquietud, no dude en contactarnos.

Saludos cordiales,
SEPAS

---
Este es un correo automático. Por favor, no responda a este mensaje.
"""
    
    def _crear_csv_emails(self, borradores: List[Dict]):
        """
        Crea un archivo CSV con los datos de los emails para importación
        
        Args:
            borradores: Lista de borradores de email
        """
        import csv
        
        archivo_csv = self.carpeta_consentimientos / "emails_lista.csv"
        
        with open(archivo_csv, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Email', 'Nombre', 'Asunto', 'Archivo Adjunto'])
            
            for borrador in borradores:
                writer.writerow([
                    borrador['email'],
                    borrador['nombre'],
                    borrador['asunto'],
                    Path(borrador['archivo']).name if borrador['archivo'] else ''
                ])
