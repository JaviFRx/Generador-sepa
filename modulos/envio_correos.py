"""Envío directo de PDFs verificados y comprobación de la copia en Enviados."""

import csv
import re
import smtplib
import ssl
import time
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import make_msgid
from pathlib import Path

from modulos.envio_borradores import EnviadorBorradores, HistorialBorradores, bloquear_lote, huella
from modulos.preparador_emails import PreparadorEmails


class SMTPConRegistro(smtplib.SMTP_SSL):
    """Conserva solo la respuesta a DATA, nunca la autenticación."""

    def data(self, mensaje):
        respuesta = super().data(mensaje)
        self.respuesta_data = respuesta
        return respuesta


class EnviadorCorreos:
    def __init__(self, carpeta):
        self.carpeta = Path(carpeta)

    def _mensajes(self, config):
        return PreparadorEmails(self.carpeta)._preparar_mensajes(
            config['ruta_excel'], config['gmail_usuario'].strip(), config.get('nombre_remitente'),
            config.get('asunto_base'), config.get('cuerpo_personalizado'),
        )

    @staticmethod
    def _esperado(mensaje):
        adjunto = next(mensaje.iter_attachments())
        return {'archivo': adjunto.get_filename(), 'email': mensaje['To'].addresses[0].addr_spec,
                'sha256': huella(adjunto.get_payload(decode=True))}

    def preparar_envio(self, config, progreso=None):
        with bloquear_lote(self.carpeta):
            mensajes = self._mensajes(config)
            cuenta = config['gmail_usuario'].strip()
            historial = HistorialBorradores(self.carpeta, cuenta)
            plan = {'cuenta': cuenta, 'mensajes': [], 'omitidos': 0}
            for i, mensaje in enumerate(mensajes, 1):
                esperado = self._esperado(mensaje)
                entrada = historial.registros.get(esperado['archivo'])
                if entrada:
                    if (entrada['email'].casefold() != esperado['email'].casefold()
                            or entrada['sha256'] != esperado['sha256']):
                        raise ValueError('El historial no coincide con el destinatario y PDF del lote.')
                    if entrada['estado'] == 'enviado':
                        plan['omitidos'] += 1
                        continue
                    if entrada['estado'] != 'borrador':
                        raise ValueError('Hay un envío sin confirmar. Revisa Gmail antes de continuar; no se repetirá.')
                # Nunca reutilizar la identidad de un borrador previamente guardado en Gmail.
                mensaje.replace_header('Message-ID', make_msgid(domain=cuenta.rsplit('@', 1)[-1]))
                EnviadorBorradores._validar_mensaje(mensaje, esperado, cuenta)
                contenido = mensaje.as_bytes()
                plan['mensajes'].append(dict(esperado, contenido=contenido, huella=huella(contenido),
                                            asunto=str(mensaje['Subject'])))
                if progreso:
                    progreso(i, len(mensajes))
            return plan

    @staticmethod
    def _carpeta_enviados(servidor):
        estado, carpetas = servidor.list()
        if estado != 'OK':
            raise ValueError('No se pueden consultar las carpetas de Gmail.')
        encontradas = []
        for linea in carpetas or []:
            partes = re.fullmatch(rb'\(([^)]*)\)\s+(?:"(?:[^"\\]|\\.)*"|NIL)\s+(.+)', linea) if isinstance(linea, bytes) else None
            if partes and b'\\sent' in partes[1].lower().split():
                encontradas.append(partes[2])
        if len(encontradas) != 1:
            raise ValueError('No se ha identificado una única carpeta Enviados en Gmail.')
        return encontradas[0]

    @staticmethod
    def _verificar_enviados(servidor, carpeta, mensaje, esperado, cuenta):
        for intento in range(3):
            if intento:
                time.sleep(intento * 2)
            estado, _ = servidor.select(carpeta, readonly=True)
            if estado != 'OK':
                raise ValueError('No se puede consultar Enviados.')
            estado, datos = servidor.uid('SEARCH', None, 'HEADER', 'Message-ID', f'"{mensaje["Message-ID"]}"')
            if estado != 'OK':
                raise ValueError('Gmail no pudo buscar la copia enviada.')
            uids = b' '.join(d for d in (datos or []) if isinstance(d, bytes)).split()
            for uid in uids:
                contenido = EnviadorBorradores._leer(servidor, uid)
                copia = BytesParser(policy=policy.default).parsebytes(contenido)
                if str(copia.get('Message-ID', '')) != str(mensaje['Message-ID']):
                    continue
                EnviadorBorradores._validar_mensaje(copia, esperado, cuenta)
                if str(copia.get('Subject', '')) != str(mensaje['Subject']):
                    raise ValueError('El asunto de la copia enviada no coincide.')
                return uid.decode('ascii')
        raise ValueError('La copia aún no se ha localizado en Enviados.')

    def _informe(self, historial):
        ruta = self.carpeta / 'informe_envios.csv'
        with ruta.open('w', encoding='utf-8-sig', newline='') as archivo:
            writer = csv.writer(archivo, delimiter=';')
            writer.writerow(['Destinatario', 'PDF', 'Estado local', 'Aceptado por Gmail (UTC)',
                             'Respuesta SMTP', 'Comprobado en Enviados (UTC)', 'Message-ID'])
            for nombre, entrada in historial.registros.items():
                writer.writerow([entrada['email'], nombre, entrada['estado'], entrada.get('aceptado_utc', ''),
                                 entrada.get('smtp_respuesta', ''), entrada.get('enviados_utc', ''),
                                 entrada.get('message_id_envio', entrada['message_id'])])

    def enviar(self, plan, config, progreso=None):
        with bloquear_lote(self.carpeta):
            cuenta = config['gmail_usuario'].strip()
            if cuenta.casefold() != plan['cuenta'].casefold():
                raise ValueError('La cuenta ha cambiado desde la confirmación.')
            # Revalidar todo el Excel y los PDFs antes de abrir conexiones.
            actuales = {e['archivo']: e for e in map(self._esperado, self._mensajes(config))}
            historial = HistorialBorradores(self.carpeta, cuenta)
            vistos = set()
            preparados = []
            for item in plan['mensajes']:
                nombre = item['archivo']
                if nombre in vistos or nombre not in actuales:
                    raise ValueError('El lote confirmado contiene PDFs duplicados o desconocidos.')
                vistos.add(nombre)
                esperado = actuales[nombre]
                if item['email'].casefold() != esperado['email'].casefold() or huella(item['contenido']) != item['huella']:
                    raise ValueError('El correo confirmado ha cambiado.')
                mensaje = BytesParser(policy=policy.default).parsebytes(item['contenido'])
                EnviadorBorradores._validar_mensaje(mensaje, esperado, cuenta)
                mids = mensaje.get_all('Message-ID', [])
                if len(mids) != 1 or not re.fullmatch(r'<[^\s<>"\\]+@[^\s<>"\\]+>', str(mids[0])):
                    raise ValueError('El identificador del correo no es válido.')
                entrada = historial.registros.get(nombre)
                if entrada and entrada['estado'] != 'borrador':
                    raise ValueError('Este correo ya se envió o tiene un envío sin confirmar. No se repetirá.')
                if entrada and (entrada['email'].casefold() != esperado['email'].casefold()
                                or entrada['sha256'] != esperado['sha256']):
                    raise ValueError('El historial ha cambiado desde la confirmación.')
                preparados.append((item, mensaje, esperado))
            resultado = {'enviados': 0, 'verificados': 0, 'omitidos': plan['omitidos'],
                         'pendientes': len(preparados), 'errores': [], 'avisos': []}
            if not preparados:
                return resultado
            try:
                with EnviadorBorradores._conectar(config) as imap:
                    carpeta = self._carpeta_enviados(imap)
                    estado, _ = imap.select(carpeta, readonly=True)
                    if estado != 'OK':
                        raise ValueError('No se puede abrir Enviados para comprobar los envíos.')
                    with SMTPConRegistro('smtp.gmail.com', 465, context=ssl.create_default_context(), timeout=30) as smtp:
                        smtp.login(cuenta, config['gmail_app_password'])
                        for i, (item, mensaje, esperado) in enumerate(preparados, 1):
                            nombre = item['archivo']
                            try:
                                historial.registrar(mensaje)
                                entrada = historial.registros[nombre]
                                entrada.update(estado='sin_confirmar', message_id_envio=str(mensaje['Message-ID']),
                                               huella_correo=item['huella'])
                                historial.guardar()
                                smtp.respuesta_data = None
                                rechazados = smtp.sendmail(cuenta, [item['email']], item['contenido'])
                                if rechazados:
                                    raise ValueError('Gmail no aceptó al destinatario.')
                                codigo, respuesta = smtp.respuesta_data or (None, b'')
                                if codigo != 250:
                                    raise ValueError('No se pudo registrar la aceptación de Gmail.')
                                entrada.update(estado='enviado', aceptado_utc=datetime.now(timezone.utc).isoformat(),
                                               smtp_codigo=codigo, smtp_respuesta=respuesta.decode('utf-8', errors='replace'))
                                resultado['enviados'] += 1
                                resultado['pendientes'] -= 1
                                historial.guardar()
                            except Exception as exc:
                                resultado['errores'].append(f"{item['email']}: {str(exc)[:200]}. Revisa Gmail; no se reintentará automáticamente.")
                                break
                            try:
                                uid = self._verificar_enviados(imap, carpeta, mensaje, esperado, cuenta)
                                entrada.update(enviados_utc=datetime.now(timezone.utc).isoformat(), uid_enviados=uid)
                                historial.guardar()
                                resultado['verificados'] += 1
                            except Exception as exc:
                                resultado['avisos'].append(f"{item['email']}: aceptado por Gmail, pero no comprobado en Enviados. "
                                                           f"{str(exc)[:200]} Se detiene el lote; este correo no se repetirá.")
                                break
                            if progreso:
                                progreso(i, len(preparados))
            except Exception as exc:
                if not resultado['enviados'] and not resultado['errores']:
                    raise
                resultado['avisos'].append(f"La conexión terminó con un error: {str(exc)[:200]}. "
                                           "Se conservan los estados registrados; no repitas los correos aceptados.")
            try:
                self._informe(historial)
            except OSError:
                resultado['avisos'].append('No se pudo escribir informe_envios.csv. El historial JSON conserva los registros.')
            return resultado
