"""Envío del contenido recuperado de Gmail, con comprobación y registro local."""

import hashlib
import imaplib
import json
import os
import smtplib
import ssl
from email import policy
from email.parser import BytesParser
from pathlib import Path
from contextlib import contextmanager


@contextmanager
def bloquear_lote(carpeta):
    ruta = Path(carpeta) / "operacion_correo.lock"
    try:
        descriptor = os.open(ruta, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        raise ValueError("Hay otra operación de correo abierta para esta carpeta. "
                         "Si el programa se cerró inesperadamente, revisa Gmail y el historial "
                         "antes de retirar operacion_correo.lock.") from exc
    try:
        os.close(descriptor)
        yield
    finally:
        ruta.unlink()


def huella(contenido):
    return hashlib.sha256(contenido).hexdigest()


class HistorialBorradores:
    def __init__(self, carpeta, cuenta):
        self.ruta = Path(carpeta) / "historial_borradores.json"
        self.cuenta = cuenta.strip().casefold()
        if self.ruta.exists():
            try:
                self.datos = json.loads(self.ruta.read_text(encoding="utf-8"))
                if self.datos["version"] != 1 or not isinstance(self.datos["cuentas"], dict):
                    raise ValueError("Formato desconocido")
            except (ValueError, KeyError, TypeError) as exc:
                raise ValueError("No se puede leer el historial de envíos. No se enviará ningún correo.") from exc
        else:
            self.datos = {"version": 1, "cuentas": {}}
        self.registros = self.datos["cuentas"].setdefault(self.cuenta, {})

    def guardar(self):
        temporal = self.ruta.with_suffix(".tmp")
        temporal.write_text(json.dumps(self.datos, ensure_ascii=False, indent=2), encoding="utf-8")
        temporal.replace(self.ruta)

    def comprobar_creacion(self, mensajes):
        for mensaje in mensajes:
            nombre = next(mensaje.iter_attachments()).get_filename()
            anterior = self.registros.get(nombre)
            if anterior and anterior["estado"] != "borrador":
                raise ValueError("Este lote ya tiene envíos registrados o pendientes de confirmar. "
                                 "Revisa Gmail antes de crear nuevos borradores.")

    def registrar(self, mensaje):
        adjunto = next(mensaje.iter_attachments())
        nombre = adjunto.get_filename()
        if nombre not in self.registros:
            self.registros[nombre] = {
                "email": str(mensaje["To"]), "sha256": huella(adjunto.get_payload(decode=True)),
                "message_id": str(mensaje["Message-ID"]), "estado": "borrador",
            }
            self.guardar()

    def marcar(self, nombre, estado):
        self.registros[nombre]["estado"] = estado
        self.guardar()


class EnviadorBorradores:
    def __init__(self, carpeta):
        self.carpeta = Path(carpeta)

    @staticmethod
    def _conectar(configuracion):
        servidor = imaplib.IMAP4_SSL("imap.gmail.com", 993,
                                    ssl_context=ssl.create_default_context(), timeout=30)
        try:
            servidor.login(configuracion["gmail_usuario"], configuracion["gmail_app_password"])
        except Exception:
            try:
                servidor.logout()
            except Exception:
                pass
            raise
        return servidor

    @staticmethod
    def _comprobar_retirada_individual(servidor):
        # imaplib conserva en capabilities la respuesta anterior al login.
        # Gmail anuncia las extensiones de la sesión autenticada al consultarlas de nuevo.
        estado, datos = servidor.capability()
        if estado != "OK":
            raise ValueError("No se pudieron consultar las funciones de Gmail. No se ha enviado ningún correo.")
        capacidades = set()
        for linea in datos or []:
            if isinstance(linea, bytes):
                capacidades.update(linea.upper().split())
        if b"UIDPLUS" not in capacidades:
            raise ValueError("Gmail no ha confirmado que permita retirar individualmente los borradores. "
                             "No se ha enviado ningún correo.")

    @staticmethod
    def _seleccionar(servidor, carpeta, escritura=False):
        estado, _ = servidor.select(carpeta, readonly=not escritura)
        if estado != "OK":
            raise ValueError("No se puede abrir Borradores en Gmail.")
        _, valores = servidor.response("UIDVALIDITY")
        if not valores or not isinstance(valores[0], bytes) or not valores[0].isdigit():
            raise ValueError("Gmail no ha confirmado la identidad de la carpeta Borradores.")
        return valores[0]

    @staticmethod
    def _leer(servidor, uid):
        estado, datos = servidor.uid("FETCH", uid, "(BODY.PEEK[])")
        contenidos = [d[1] for d in (datos or []) if isinstance(d, tuple) and isinstance(d[1], bytes)]
        if estado != "OK" or len(contenidos) != 1:
            raise ValueError("Un borrador ha desaparecido o no se puede leer. Vuelve a revisar el lote.")
        return contenidos[0]

    @staticmethod
    def _validar_mensaje(mensaje, esperado, cuenta):
        if mensaje.defects:
            raise ValueError("Un borrador tiene un formato de correo inválido.")
        for cabecera in ("To", "From"):
            valores = mensaje.get_all(cabecera, [])
            if len(valores) != 1 or valores[0].defects or len(valores[0].addresses) != 1:
                raise ValueError(f"Cabecera {cabecera} ambigua en un borrador.")
        if mensaje["To"].addresses[0].addr_spec.casefold() != esperado["email"].casefold():
            raise ValueError(f"El destinatario del PDF {esperado['archivo']} ha cambiado en Gmail.")
        if mensaje["From"].addresses[0].addr_spec.casefold() != cuenta.casefold():
            raise ValueError("El remitente del borrador no coincide con la cuenta configurada.")
        if any(mensaje.get_all(h) for h in ("Cc", "Bcc", "Resent-To", "Resent-Cc", "Resent-Bcc", "Resent-From")):
            raise ValueError("El borrador contiene destinatarios adicionales. Revisa Gmail antes de enviar.")
        adjuntos = list(mensaje.iter_attachments())
        if len(adjuntos) != 1:
            raise ValueError("Cada borrador debe contener únicamente su PDF verificado.")
        adjunto = adjuntos[0]
        if (adjunto.get_filename() != esperado["archivo"]
                or huella(adjunto.get_payload(decode=True) or b"") != esperado["sha256"]):
            raise ValueError(f"El PDF de {esperado['email']} no coincide con el generado para esa persona.")

    def preparar_envio(self, configuracion, progreso=None):
        with bloquear_lote(self.carpeta):
            return self._preparar_envio(configuracion, progreso)

    def _preparar_envio(self, configuracion, progreso=None):
        """Descarga y comprueba el lote; no abre ninguna conexión SMTP."""
        from modulos.preparador_emails import PreparadorEmails
        preparador = PreparadorEmails(self.carpeta)
        cuenta = configuracion["gmail_usuario"].strip()
        originales = preparador._preparar_mensajes(
            configuracion["ruta_excel"], cuenta, configuracion.get("nombre_remitente"),
            configuracion.get("asunto_base"), configuracion.get("cuerpo_personalizado"),
        )
        historial = HistorialBorradores(self.carpeta, cuenta)
        esperados = {}
        omitidos = 0
        for original in originales:
            adjunto = next(original.iter_attachments())
            nombre = adjunto.get_filename()
            entrada = historial.registros.get(nombre)
            if entrada and entrada["estado"] == "enviado":
                omitidos += 1
                continue
            if entrada and entrada["estado"] != "borrador":
                raise ValueError("Hay un envío sin confirmar. Comprueba Enviados en Gmail; "
                                 "la aplicación no lo repetirá automáticamente.")
            esperados[nombre] = {"archivo": nombre, "email": original["To"].addresses[0].addr_spec,
                                 "sha256": huella(adjunto.get_payload(decode=True)),
                                 "message_id": entrada["message_id"] if entrada else str(original["Message-ID"]),
                                 "registrado": bool(entrada), "original": original}
            if entrada and (entrada["sha256"] != esperados[nombre]["sha256"]
                            or entrada["email"].casefold() != esperados[nombre]["email"].casefold()):
                raise ValueError("El historial no coincide con los destinatarios y PDFs del lote.")
        plan = {"cuenta": cuenta, "mensajes": [], "omitidos": omitidos}
        if not esperados:
            return plan
        with self._conectar(configuracion) as servidor:
            carpeta = preparador._carpeta_borradores(servidor)
            plan["carpeta"] = carpeta
            plan["uidvalidity"] = self._seleccionar(servidor, carpeta)
            estado, datos = servidor.uid("SEARCH", None, "UNDELETED")
            if estado != "OK":
                raise ValueError("No se pudieron consultar los borradores de Gmail.")
            encontrados = {}
            uids = b" ".join(d for d in (datos or []) if isinstance(d, bytes)).split()
            for i, uid in enumerate(uids, 1):
                if not uid.isdigit():
                    raise ValueError("Gmail devolvió un identificador de borrador inválido.")
                contenido = self._leer(servidor, uid)
                mensaje = BytesParser(policy=policy.default).parsebytes(contenido)
                nombres = {p.get_filename() for p in mensaje.iter_attachments()}
                candidatos = [e for e in esperados.values()
                              if str(mensaje.get("Message-ID", "")) == e["message_id"]
                              or (e["registrado"] and e["archivo"] in nombres)]
                if candidatos:
                    if len(candidatos) != 1:
                        raise ValueError("Un borrador contiene PDFs de varias personas.")
                    esperado = candidatos[0]
                    self._validar_mensaje(mensaje, esperado, cuenta)
                    if esperado["archivo"] in encontrados:
                        raise ValueError("Hay varios borradores para el mismo PDF. Elimina el duplicado en Gmail.")
                    encontrados[esperado["archivo"]] = {
                        "uid": uid, "contenido": contenido, "huella": huella(contenido),
                        "archivo": esperado["archivo"], "email": esperado["email"],
                        "asunto": str(mensaje.get("Subject", "")),
                    }
                if progreso:
                    progreso(i, len(uids))
            if len(encontrados) != len(esperados):
                raise ValueError("Faltan borradores identificables del lote en Gmail. "
                                 "Comprueba si los has eliminado o enviado manualmente. "
                                 "Para borradores de versiones anteriores, pulsa Crear borradores "
                                 "con el mismo asunto y texto para registrar los existentes.")
            for nombre, esperado in esperados.items():
                historial.registrar(esperado["original"])
                plan["mensajes"].append(encontrados[nombre])
        return plan

    def enviar(self, plan, configuracion, progreso=None):
        with bloquear_lote(self.carpeta):
            return self._enviar(plan, configuracion, progreso)

    def _enviar(self, plan, configuracion, progreso=None):
        """Envía exclusivamente el contenido descargado y confirmado por el usuario."""
        cuenta = configuracion["gmail_usuario"].strip()
        if plan["cuenta"].casefold() != cuenta.casefold():
            raise ValueError("La cuenta ha cambiado desde la revisión.")
        resultado = {"enviados": 0, "omitidos": plan["omitidos"], "pendientes": len(plan["mensajes"]),
                     "errores": [], "avisos": []}
        if not plan["mensajes"]:
            return resultado
        historial = HistorialBorradores(self.carpeta, cuenta)
        for item in plan["mensajes"]:
            if historial.registros[item["archivo"]]["estado"] != "borrador":
                raise ValueError("Este lote ya se ha enviado o tiene envíos sin confirmar. No se repetirá.")
            if huella(item["contenido"]) != item["huella"]:
                raise ValueError("El contenido confirmado ha cambiado. Vuelve a revisar los borradores.")
            esperado = dict(historial.registros[item["archivo"]], archivo=item["archivo"])
            self._validar_mensaje(BytesParser(policy=policy.default).parsebytes(item["contenido"]), esperado, cuenta)
            if item["email"].casefold() != esperado["email"].casefold():
                raise ValueError("El destinatario confirmado ha cambiado.")
        with self._conectar(configuracion) as servidor:
            self._comprobar_retirada_individual(servidor)
            if self._seleccionar(servidor, plan["carpeta"], escritura=True) != plan["uidvalidity"]:
                raise ValueError("La carpeta de Gmail ha cambiado. Vuelve a revisar los borradores.")
            # Validar TODO el lote antes de enviar el primer correo.
            for item in plan["mensajes"]:
                if huella(self._leer(servidor, item["uid"])) != item["huella"]:
                    raise ValueError("Un borrador ha cambiado desde la confirmación. Vuelve a revisar el lote.")
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=30) as smtp:
                smtp.login(cuenta, configuracion["gmail_app_password"])
                for i, item in enumerate(plan["mensajes"], 1):
                    try:
                        if huella(self._leer(servidor, item["uid"])) != item["huella"]:
                            raise ValueError("El borrador ha cambiado en Gmail durante el envío del lote")
                        # Una interrupción deja este estado persistido y bloquea la repetición.
                        historial.marcar(item["archivo"], "sin_confirmar")
                        rechazados = smtp.sendmail(cuenta, [item["email"]], item["contenido"])
                        if rechazados:
                            raise ValueError("El servidor no aceptó el destinatario.")
                        historial.marcar(item["archivo"], "enviado")
                        resultado["enviados"] += 1
                        resultado["pendientes"] -= 1
                    except Exception as exc:
                        resultado["errores"].append(f"{item['email']}: {str(exc)[:200]}. "
                                                    "Revisa Enviados en Gmail antes de volver a intentarlo.")
                        break
                    try:
                        # No retirar un borrador que se haya editado mientras se enviaba.
                        if huella(self._leer(servidor, item["uid"])) != item["huella"]:
                            raise ValueError("el borrador ha cambiado")
                        estado, _ = servidor.uid("STORE", item["uid"], "+FLAGS.SILENT", r"(\Deleted)")
                        if estado != "OK":
                            raise ValueError("no se pudo marcar el borrador enviado")
                        estado, _ = servidor.uid("EXPUNGE", item["uid"])
                        if estado != "OK":
                            raise ValueError("no se pudo retirar el borrador enviado")
                    except Exception as exc:
                        resultado["avisos"].append(f"{item['email']} ya está enviado, pero su borrador "
                                                   f"requiere revisión en Gmail: {str(exc)[:150]}.")
                        break
                    if progreso:
                        progreso(i, len(plan["mensajes"]))
        return resultado
