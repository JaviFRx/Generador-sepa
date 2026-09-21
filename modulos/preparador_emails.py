"""
Módulo para preparar emails con consentimientos adjuntos
"""

from pathlib import Path
from typing import List, Dict, Optional
import re
import imaplib
import hashlib
import json
import ssl
from email.message import EmailMessage
from email import policy
from email.utils import formatdate, formataddr
from modulos.registro_consentimientos import RegistroConsentimientos


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
        
        archivos_generados = RegistroConsentimientos(self.carpeta_consentimientos).validar(datos)
        emails = [self._encontrar_email(registro) for registro in datos]
        
        # Crear borradores
        borradores = []
        archivo_salida = self.carpeta_consentimientos / "emails_para_enviar.txt"
        
        with open(archivo_salida, 'w', encoding='utf-8') as f:
            f.write("=" * 80 + "\n")
            f.write("REVISIÓN DE CORREOS Y ADJUNTOS (PDF)\n")
            f.write(f"Generado: {Path(ruta_excel).name}\n")
            f.write("=" * 80 + "\n\n")
            
            for i, registro in enumerate(datos, 1):
                # Buscar email en el registro
                email = emails[i - 1]
                nombre = self._encontrar_nombre(registro)
                
                # Buscar archivo de consentimiento correspondiente
                archivo_consentimiento = archivos_generados[i - 1]["archivo"]
                
                # Generar asunto
                asunto = asunto_base or "Consentimiento"
                
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

    def crear_borradores_gmail(
        self, ruta_excel: str, gmail_usuario: str, gmail_app_password: str,
        nombre_remitente: Optional[str] = None, asunto_base: str = "Consentimiento",
        cuerpo_personalizado: str = None, progreso=None,
    ) -> Dict:
        from modulos.envio_borradores import bloquear_lote
        with bloquear_lote(self.carpeta_consentimientos):
            return self._crear_borradores_gmail(
                ruta_excel, gmail_usuario, gmail_app_password, nombre_remitente,
                asunto_base, cuerpo_personalizado, progreso,
            )

    def _crear_borradores_gmail(
        self, ruta_excel, gmail_usuario, gmail_app_password, nombre_remitente,
        asunto_base, cuerpo_personalizado, progreso,
    ):
        """Guarda mensajes con adjunto en Borradores mediante IMAP; nunca los envía."""
        if not gmail_usuario or not gmail_app_password:
            raise ValueError("Indica la cuenta Gmail y su contraseña de aplicación.")
        mensajes = self._preparar_mensajes(
            ruta_excel, gmail_usuario, nombre_remitente, asunto_base, cuerpo_personalizado
        )
        from modulos.envio_borradores import HistorialBorradores
        historial = HistorialBorradores(self.carpeta_consentimientos, gmail_usuario)
        historial.comprobar_creacion(mensajes)
        # Serializar el lote completo antes de realizar cualquier escritura remota.
        contenidos = [mensaje.as_bytes() for mensaje in mensajes]
        resultado = {"total": len(mensajes), "creados": 0, "existentes": 0,
                     "fallidos": [], "pendientes": len(mensajes)}
        if progreso:
            progreso(0, len(mensajes), 0, 0)
        with imaplib.IMAP4_SSL("imap.gmail.com", 993, ssl_context=ssl.create_default_context(),
                             timeout=30) as servidor:
            servidor.login(gmail_usuario, gmail_app_password)
            carpeta = self._carpeta_borradores(servidor)
            estado, _ = servidor.select(carpeta, readonly=True)
            if estado != "OK":
                raise ValueError("Gmail no permite consultar su carpeta Borradores.")
            for i, (mensaje, contenido) in enumerate(zip(mensajes, contenidos), 1):
                try:
                    # Un reintento no duplica un borrador que ya esté en esa carpeta.
                    estado, coincidencias = servidor.uid(
                        "SEARCH", None, "HEADER", "Message-ID", f'"{mensaje["Message-ID"]}"'
                    )
                    if estado != "OK":
                        raise ValueError("No se pudo comprobar si el borrador ya existe.")
                    if any(dato.strip() for dato in (coincidencias or []) if isinstance(dato, bytes)):
                        resultado["existentes"] += 1
                    else:
                        estado, _ = servidor.append(carpeta, r"(\Draft)", None, contenido)
                        if estado != "OK":
                            raise ValueError("Gmail no confirmó la creación del borrador.")
                        resultado["creados"] += 1
                    historial.registrar(mensaje)
                    resultado["pendientes"] = len(mensajes) - i
                    if progreso:
                        progreso(i, len(mensajes), resultado["creados"], resultado["existentes"])
                except (imaplib.IMAP4.error, OSError, ValueError) as exc:
                    resultado["fallidos"].append({
                        "numero": i, "email": str(mensaje["To"]),
                        "motivo": f"{str(exc)[:200]} Revisa Borradores en Gmail antes de reintentar.",
                    })
                    resultado["pendientes"] = len(mensajes) - i
                    # Una pérdida de conexión puede ocurrir después de guardar el mensaje.
                    # Se detiene el lote sin reintentos automáticos.
                    break
        return resultado

    @staticmethod
    def _carpeta_borradores(servidor):
        """Detecta la carpeta por su atributo IMAP, independientemente del idioma."""
        estado, carpetas = servidor.list()
        if estado != "OK":
            raise ValueError("No se pudieron consultar las carpetas de Gmail.")
        encontradas = []
        for linea in carpetas or []:
            if not isinstance(linea, bytes):
                continue
            partes = re.fullmatch(rb'\(([^)]*)\)\s+(?:"(?:[^"\\]|\\.)*"|NIL)\s+(.+)', linea)
            if partes and b"\\drafts" in partes[1].lower().split():
                encontradas.append(partes[2])
        if len(encontradas) != 1:
            raise ValueError("No se ha identificado una única carpeta Borradores en Gmail.")
        return encontradas[0]

    def _preparar_mensajes(self, ruta_excel, gmail_usuario, nombre_remitente=None,
                           asunto_base="Consentimiento", cuerpo_personalizado=None):
        from modulos.lector_excel import LectorExcel

        lector = LectorExcel(ruta_excel)
        datos = lector.leer_datos()

        # Validar y cargar todos los adjuntos antes de conectar con Gmail.
        # El borrador conserva los mismos bytes verificados.
        archivos_generados = RegistroConsentimientos(self.carpeta_consentimientos).validar(datos)
        mensajes = []
        for i, registro in enumerate(datos, 1):
            email_destino = self._encontrar_email(registro)
            nombre = self._encontrar_nombre(registro)
            adjunto = archivos_generados[i - 1]
            cuerpo = (cuerpo_personalizado.replace("{nombre}", nombre)
                      if cuerpo_personalizado else self._generar_cuerpo_email(nombre))
            mensaje = EmailMessage(policy=policy.SMTP)
            mensaje["From"] = formataddr((nombre_remitente or "", gmail_usuario))
            mensaje["To"] = email_destino
            mensaje["Subject"] = asunto_base or "Consentimiento"
            mensaje.set_content(cuerpo)
            mensaje.add_attachment(
                adjunto["contenido"], maintype="application", subtype="pdf",
                filename=Path(adjunto["archivo"]).name,
            )
            identidad = json.dumps({
                "remitente": str(mensaje["From"]), "destino": email_destino,
                "asunto": str(mensaje["Subject"]), "cuerpo": cuerpo,
                "archivo": Path(adjunto["archivo"]).name,
                "sha256": hashlib.sha256(adjunto["contenido"]).hexdigest(),
            }, sort_keys=True, ensure_ascii=False).encode("utf-8")
            mensaje["Message-ID"] = f"<sepas-{hashlib.sha256(identidad).hexdigest()}@borradores.local>"
            mensaje["Date"] = formatdate(localtime=True)
            mensajes.append(mensaje)

        return mensajes

    def _encontrar_email(self, registro: Dict) -> str:
        """Exige un único destinatario inequívoco, sin direcciones inventadas."""
        candidatos = set()
        for clave, valor in registro.items():
            if 'email' in clave.lower() or 'correo' in clave.lower() or 'mail' in clave.lower():
                email = str(valor or "").strip()
                if not email:
                    continue
                if not re.fullmatch(r'[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+', email):
                    raise ValueError("Borradores bloqueados: hay un email inválido o varios destinatarios en una celda.")
                candidatos.add(email)
        if len(candidatos) != 1:
            raise ValueError("Borradores bloqueados: falta un email único por registro; revisa las columnas de correo.")
        return candidatos.pop()
    
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
