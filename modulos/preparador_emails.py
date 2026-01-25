"""
Módulo para preparar emails con consentimientos adjuntos
"""

from pathlib import Path
from typing import List, Dict, Optional
import openpyxl
import mimetypes
import smtplib
import ssl
from email.message import EmailMessage


class PreparadorEmails:
    """Clase para preparar emails con consentimientos"""
    
    def __init__(self, carpeta_consentimientos: str):
        """
        Inicializa el preparador de emails
        
        Args:
            carpeta_consentimientos: Carpeta donde están los consentimientos generados
        """
        self.carpeta_consentimientos = Path(carpeta_consentimientos)
        
    def crear_borradores_email(self, ruta_excel: str, asunto_base: str = "Consentimiento", 
                               cuerpo_personalizado: str = None) -> List[Dict]:
        """
        Crea borradores de email basados en el Excel y los consentimientos generados
        
        Args:
            ruta_excel: Ruta al archivo Excel con los datos
            asunto_base: Asunto base para los emails
            cuerpo_personalizado: Cuerpo personalizado del email (si None, usa el por defecto)
            
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
                
                # Generar asunto
                asunto = f"{asunto_base} - {nombre}" if nombre else asunto_base
                
                # Generar cuerpo
                if cuerpo_personalizado:
                    cuerpo = cuerpo_personalizado.replace("{nombre}", nombre)
                else:
                    cuerpo = self._generar_cuerpo_email(nombre)
                
                # Crear borrador
                borrador = {
                    'numero': i,
                    'email': email,
                    'nombre': nombre,
                    'archivo': archivo_consentimiento,
                    'asunto': asunto,
                    'cuerpo': cuerpo
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

    def enviar_emails_gmail(
        self,
        ruta_excel: str,
        gmail_usuario: str,
        gmail_app_password: str,
        nombre_remitente: Optional[str] = None,
        asunto_base: str = "Consentimiento",
        cuerpo_personalizado: str = None,
    ) -> Dict:
        """
        Envía emails reales usando Gmail SMTP con los consentimientos adjuntos.

        Requiere contraseña de aplicación (App Password) de Gmail.

        Args:
            ruta_excel: Ruta al archivo Excel con los datos
            gmail_usuario: Cuenta Gmail desde la que se enviarán los correos
            gmail_app_password: Contraseña de aplicación de Gmail
            nombre_remitente: Nombre visible del remitente (opcional)
            asunto_base: Texto base del asunto
            cuerpo_personalizado: Cuerpo personalizado del email (si None, usa el por defecto)

        Returns:
            Diccionario con resumen de enviados y fallidos
        """
        if not gmail_usuario or not gmail_app_password:
            raise Exception("Debe indicar usuario Gmail y contraseña de aplicación")

        from modulos.lector_excel import LectorExcel

        lector = LectorExcel(ruta_excel)
        datos = lector.leer_datos()

        archivos_generados = list(self.carpeta_consentimientos.glob("*.pdf"))

        enviados = 0
        fallidos = []

        contexto_ssl = ssl.create_default_context()

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto_ssl) as server:
            server.login(gmail_usuario, gmail_app_password)

            for i, registro in enumerate(datos, 1):
                email_destino = self._encontrar_email(registro)
                nombre = self._encontrar_nombre(registro)
                archivo_consentimiento = self._encontrar_consentimiento(registro, archivos_generados)

                if not email_destino or "@" not in email_destino:
                    fallidos.append({"numero": i, "email": email_destino, "motivo": "Email inválido"})
                    continue

                if not archivo_consentimiento:
                    fallidos.append({"numero": i, "email": email_destino, "motivo": "PDF no encontrado"})
                    continue

                asunto = f"{asunto_base}" if asunto_base else "Consentimiento"
                
                # Generar cuerpo personalizado
                if cuerpo_personalizado:
                    cuerpo = cuerpo_personalizado.replace("{nombre}", nombre)
                else:
                    cuerpo = self._generar_cuerpo_email(nombre)

                mensaje = EmailMessage()
                if nombre_remitente:
                    mensaje["From"] = f"{nombre_remitente} <{gmail_usuario}>"
                else:
                    mensaje["From"] = gmail_usuario
                mensaje["To"] = email_destino
                mensaje["Subject"] = asunto
                mensaje.set_content(cuerpo)

                tipo_mime, _ = mimetypes.guess_type(archivo_consentimiento)
                if not tipo_mime:
                    tipo_mime = "application/pdf"
                tipo_main, tipo_sub = tipo_mime.split("/", 1)

                with open(archivo_consentimiento, "rb") as f:
                    mensaje.add_attachment(
                        f.read(),
                        maintype=tipo_main,
                        subtype=tipo_sub,
                        filename=Path(archivo_consentimiento).name,
                    )

                try:
                    server.send_message(mensaje)
                    enviados += 1
                except Exception as e:
                    fallidos.append({"numero": i, "email": email_destino, "motivo": str(e)[:200]})

        return {
            "total": len(datos),
            "enviados": enviados,
            "fallidos": fallidos,
        }
    
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
        # Si no hay archivos, retornar None
        if not archivos:
            return None
        
        # Si solo hay un archivo, retornarlo (asumiendo que es el único)
        if len(archivos) == 1:
            return str(archivos[0])
        
        # Buscar por valores de cualquier campo en el nombre del archivo
        for archivo in archivos:
            nombre_archivo = archivo.stem.lower()
            
            # Buscar cualquier valor del registro en el nombre del archivo
            for clave, valor in registro.items():
                if valor and str(valor).lower() in nombre_archivo:
                    return str(archivo)
        
        # Si no encuentra por coincidencia exacta, retornar el primer archivo
        # (como fallback cuando solo hay uno)
        if len(archivos) == 1:
            return str(archivos[0])
        
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
