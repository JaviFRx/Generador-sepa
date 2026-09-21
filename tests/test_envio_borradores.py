import tempfile
import unittest
from pathlib import Path
from email import policy
from email.parser import BytesParser
from unittest.mock import patch

from openpyxl import Workbook

from modulos.envio_borradores import EnviadorBorradores, HistorialBorradores, bloquear_lote
from modulos.preparador_emails import PreparadorEmails
from modulos.registro_consentimientos import RegistroConsentimientos


class GmailSimulado:
    # Como imaplib: la caché sigue conteniendo las capacidades previas al login.
    capabilities = (b"IMAP4REV1", b"AUTH=PLAIN")

    def __init__(self, mensajes):
        self.mensajes = dict(mensajes)
        self.uidvalidity = b"123"
        self.retirados = []
        self.fallar_retirada = False
        self.respuesta_capability = ("OK", [b"IMAP4rev1 UIDPLUS"])
        self.consultas_capability = 0

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def list(self):
        return "OK", [b'(\\Drafts) "/" "[Gmail]/Borradores"']

    def capability(self):
        self.consultas_capability += 1
        return self.respuesta_capability

    def select(self, *args, **kwargs):
        return "OK", [str(len(self.mensajes)).encode()]

    def response(self, nombre):
        return nombre, [self.uidvalidity]

    def uid(self, operacion, *args):
        if operacion == "SEARCH":
            return "OK", [b" ".join(self.mensajes)]
        uid = args[0]
        if operacion == "FETCH":
            return "OK", [(b"BODY[]", self.mensajes[uid])] if uid in self.mensajes else []
        if operacion == "STORE":
            return ("NO" if self.fallar_retirada else "OK"), []
        if operacion == "EXPUNGE":
            self.retirados.append(uid)
            self.mensajes.pop(uid)
            return "OK", []
        raise AssertionError(operacion)


class EnvioBorradoresTest(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.carpeta = Path(self.temporal.name)
        self.datos = [{"nombre": "Ana", "email": "ana@example.test"},
                      {"nombre": "Luis", "email": "luis@example.test"}]
        excel = self.carpeta / "datos.xlsx"
        libro = Workbook()
        libro.active.append(list(self.datos[0]))
        for fila in self.datos:
            libro.active.append(list(fila.values()))
        libro.save(excel)
        libro.close()
        registro = RegistroConsentimientos(self.carpeta)
        registro.iniciar(self.datos)
        for i, fila in enumerate(self.datos, 1):
            pdf = self.carpeta / f"consentimiento_{i}.pdf"
            pdf.write_bytes(b"%PDF-1.7\n" + fila["nombre"].encode())
            registro.agregar(i, fila, pdf)
        registro.finalizar()
        self.config = {"ruta_excel": str(excel), "gmail_usuario": "remitente@example.test",
                       "gmail_app_password": "clave-ficticia", "asunto_base": "SEPA",
                       "cuerpo_personalizado": "Hola {nombre}"}
        self.originales = PreparadorEmails(self.carpeta)._preparar_mensajes(
            str(excel), self.config["gmail_usuario"], asunto_base="SEPA", cuerpo_personalizado="Hola {nombre}"
        )
        historial = HistorialBorradores(self.carpeta, self.config["gmail_usuario"])
        for mensaje in self.originales:
            historial.registrar(mensaje)
        self.gmail = GmailSimulado({str(i).encode(): m.as_bytes() for i, m in enumerate(self.originales, 1)})
        self.enviador = EnviadorBorradores(self.carpeta)
        self.conexion = patch.object(EnviadorBorradores, "_conectar", return_value=self.gmail).start()
        self.smtp = patch("modulos.envio_borradores.smtplib.SMTP_SSL").start()
        self.addCleanup(patch.stopall)
        self.transporte = self.smtp.return_value.__enter__.return_value
        self.transporte.sendmail.return_value = {}

    def editar(self, uid, cambio):
        mensaje = BytesParser(policy=policy.SMTP).parsebytes(self.gmail.mensajes[uid])
        cambio(mensaje)
        self.gmail.mensajes[uid] = mensaje.as_bytes()

    def preparar(self):
        return self.enviador.preparar_envio(self.config)

    def test_envia_version_revisada_en_gmail_y_retiro_solo_esos_borradores(self):
        def cambios(mensaje):
            mensaje.replace_header("Subject", "Asunto corregido en Gmail")
            mensaje.replace_header("Message-ID", "<editado-por-gmail@example.test>")
            mensaje.get_body().set_content("Texto corregido en Gmail")
        self.editar(b"1", cambios)
        self.gmail.mensajes[b"99"] = b"From: otro@example.test\r\nTo: familiar@example.test\r\nSubject: Personal\r\n\r\nPrivado"
        plan = self.preparar()
        self.smtp.assert_not_called()
        contenido_revisado = self.gmail.mensajes[b"1"]
        self.assertEqual(plan["mensajes"][0]["asunto"], "Asunto corregido en Gmail")
        resultado = self.enviador.enviar(plan, self.config)
        self.assertEqual(resultado["enviados"], 2)
        llamada = self.transporte.sendmail.call_args_list[0]
        self.assertEqual(llamada.args, (self.config["gmail_usuario"], ["ana@example.test"], contenido_revisado))
        self.assertEqual(self.gmail.retirados, [b"1", b"2"])
        self.assertIn(b"99", self.gmail.mensajes)
        self.smtp.reset_mock()
        segundo = self.preparar()
        self.assertEqual(segundo["omitidos"], 2)
        self.assertEqual(segundo["mensajes"], [])
        self.enviador.enviar(segundo, self.config)
        self.smtp.assert_not_called()

    def test_destinatario_cambiado_bloquea_todo(self):
        self.editar(b"2", lambda m: m.replace_header("To", "otro@example.test"))
        with self.assertRaisesRegex(ValueError, "destinatario"):
            self.preparar()
        self.smtp.assert_not_called()

    def test_consulta_capacidades_autenticadas_en_vez_de_cache_previa(self):
        self.assertNotIn(b"UIDPLUS", self.gmail.capabilities)
        self.gmail.respuesta_capability = ("OK", [b"IMAP4rev1", b"uidplus X-GM-EXT-1"])
        resultado = self.enviador.enviar(self.preparar(), self.config)
        self.assertEqual(self.gmail.consultas_capability, 1)
        self.assertEqual(resultado["enviados"], 2)
        self.assertEqual(self.gmail.retirados, [b"1", b"2"])

    def test_sin_uidplus_actual_bloquea_antes_de_smtp_sin_alterar_historial(self):
        self.gmail.capabilities = (b"IMAP4REV1", b"UIDPLUS")
        self.gmail.respuesta_capability = ("OK", [b"IMAP4rev1"])
        plan = self.preparar()
        ruta = self.carpeta / "historial_borradores.json"
        anterior = ruta.read_bytes()
        with self.assertRaisesRegex(ValueError, "No se ha enviado ningún correo"):
            self.enviador.enviar(plan, self.config)
        self.smtp.assert_not_called()
        self.assertEqual(ruta.read_bytes(), anterior)
        self.assertEqual(self.gmail.retirados, [])
        self.assertEqual(len(self.gmail.mensajes), 2)

    def test_fallo_consulta_capacidades_bloquea_envio_y_permite_reintentar(self):
        plan = self.preparar()
        ruta = self.carpeta / "historial_borradores.json"
        anterior = ruta.read_bytes()
        for respuesta in (("NO", [b"UIDPLUS"]), ("OK", [None])):
            with self.subTest(respuesta=respuesta):
                self.gmail.respuesta_capability = respuesta
                with self.assertRaisesRegex(ValueError, "No se ha enviado ningún correo"):
                    self.enviador.enviar(plan, self.config)
                self.smtp.assert_not_called()
                self.assertEqual(ruta.read_bytes(), anterior)
                self.assertEqual(self.gmail.retirados, [])
        self.gmail.respuesta_capability = ("OK", [b"IMAP4rev1 UIDPLUS"])
        self.assertEqual(self.enviador.enviar(plan, self.config)["enviados"], 2)

    def test_pdf_cambiado_bloquea_todo(self):
        self.editar(b"2", lambda m: next(m.iter_attachments()).set_payload("OTRO PDF"))
        with self.assertRaisesRegex(ValueError, "PDF"):
            self.preparar()
        self.smtp.assert_not_called()

    def test_pdf_de_otra_persona_se_rechaza(self):
        primero = next(self.originales[0].iter_attachments())
        def intercambiar(mensaje):
            mensaje.set_payload([mensaje.get_body(), primero])
        self.editar(b"2", intercambiar)
        with self.assertRaises(ValueError):
            self.preparar()
        self.smtp.assert_not_called()

    def test_destinatarios_ocultos_se_rechazan(self):
        self.editar(b"2", lambda m: m.__setitem__("Bcc", "oculto@example.test"))
        with self.assertRaisesRegex(ValueError, "adicionales"):
            self.preparar()
        self.smtp.assert_not_called()

    def test_borrador_duplicado_bloquea_todo(self):
        self.gmail.mensajes[b"3"] = self.gmail.mensajes[b"1"]
        with self.assertRaisesRegex(ValueError, "varios borradores"):
            self.preparar()
        self.smtp.assert_not_called()

    def test_borrador_ausente_no_se_reconstruye(self):
        self.gmail.mensajes.pop(b"2")
        with self.assertRaisesRegex(ValueError, "Faltan borradores"):
            self.preparar()
        self.smtp.assert_not_called()

    def test_cambio_posterior_a_confirmacion_bloquea_antes_de_smtp(self):
        plan = self.preparar()
        self.editar(b"2", lambda m: m.replace_header("Subject", "Cambio posterior"))
        with self.assertRaisesRegex(ValueError, "ha cambiado"):
            self.enviador.enviar(plan, self.config)
        self.smtp.assert_not_called()

    def test_cambio_de_carpeta_bloquea_envio(self):
        plan = self.preparar()
        self.gmail.uidvalidity = b"456"
        with self.assertRaisesRegex(ValueError, "carpeta"):
            self.enviador.enviar(plan, self.config)
        self.smtp.assert_not_called()

    def test_timeout_no_repite_envio_incierto(self):
        plan = self.preparar()
        self.transporte.sendmail.side_effect = TimeoutError("Se perdió la respuesta")
        resultado = self.enviador.enviar(plan, self.config)
        self.assertEqual(resultado["enviados"], 0)
        self.assertEqual(len(resultado["errores"]), 1)
        self.assertEqual(self.gmail.retirados, [])
        self.transporte.sendmail.assert_called_once()
        with self.assertRaisesRegex(ValueError, "sin confirmar"):
            self.preparar()
        with self.assertRaisesRegex(ValueError, "sin confirmar"):
            self.enviador.enviar(plan, self.config)
        self.transporte.sendmail.assert_called_once()

    def test_fallo_al_retirar_no_permite_reenviar_lo_confirmado(self):
        plan = self.preparar()
        self.gmail.fallar_retirada = True
        resultado = self.enviador.enviar(plan, self.config)
        self.assertEqual(resultado["enviados"], 1)
        self.assertEqual(resultado["pendientes"], 1)
        self.assertEqual(len(resultado["avisos"]), 1)
        segundo = self.preparar()
        self.assertEqual(segundo["omitidos"], 1)
        self.assertEqual([m["email"] for m in segundo["mensajes"]], ["luis@example.test"])

    def test_migracion_borradores_anteriores_por_message_id(self):
        (self.carpeta / "historial_borradores.json").unlink()
        self.assertEqual(len(self.preparar()["mensajes"]), 2)
        self.assertTrue((self.carpeta / "historial_borradores.json").exists())
        self.smtp.assert_not_called()

    def test_no_envia_dos_operaciones_a_la_vez(self):
        plan = self.preparar()
        with bloquear_lote(self.carpeta):
            with self.assertRaisesRegex(ValueError, "otra operación"):
                self.enviador.enviar(plan, self.config)
        self.smtp.assert_not_called()

    def test_no_envia_si_se_altera_el_contenido_o_destino_del_plan(self):
        plan = self.preparar()
        plan["mensajes"][0]["contenido"] = b"contenido distinto"
        with self.assertRaisesRegex(ValueError, "contenido confirmado"):
            self.enviador.enviar(plan, self.config)
        self.smtp.assert_not_called()


if __name__ == "__main__":
    unittest.main()
