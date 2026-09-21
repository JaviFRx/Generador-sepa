import tempfile
import unittest
from email import policy
from email.parser import BytesParser
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from modulos.envio_borradores import HistorialBorradores
from modulos.envio_correos import EnviadorCorreos, SMTPConRegistro
from modulos.registro_consentimientos import RegistroConsentimientos


class GmailLectura:
    def __init__(self):
        self.mensajes = {}
        self.consultas = []
        self.visible = True
        self.fallar_busqueda = False
        self.retardo = 0
        self.carpetas = [b'(\\Sent) "/" "[Gmail]/Enviados"']

    def __enter__(self):
        return self

    def __exit__(self, *args):
        pass

    def list(self):
        return 'OK', self.carpetas

    def select(self, carpeta, readonly):
        assert readonly is True, 'La comprobacion no debe escribir en Gmail'
        return 'OK', [b'0']

    def uid(self, operacion, *args):
        self.consultas.append(operacion)
        if operacion == 'SEARCH':
            if self.fallar_busqueda:
                return 'NO', []
            if self.retardo:
                self.retardo -= 1
                return 'OK', [b'']
            mid = args[-1].strip('"')
            uids = [uid for uid, contenido in self.mensajes.items()
                    if str(BytesParser(policy=policy.default).parsebytes(contenido)['Message-ID']) == mid]
            return 'OK', [b' '.join(uids) if self.visible else b'']
        if operacion == 'FETCH':
            assert args[1] == '(BODY.PEEK[])'
            return 'OK', [(b'BODY[]', self.mensajes[args[0]])]
        raise AssertionError('Operacion no permitida: ' + operacion)


class EnvioCorreosTest(unittest.TestCase):
    def setUp(self):
        temporal = tempfile.TemporaryDirectory()
        self.addCleanup(temporal.cleanup)
        self.carpeta = Path(temporal.name)
        self.excel = self.carpeta / 'datos.xlsx'
        datos = [{'nombre': 'Ana', 'email': 'ana@example.test'}, {'nombre': 'Luis', 'email': 'luis@example.test'}]
        libro = Workbook()
        libro.active.append(['nombre', 'email'])
        for fila in datos:
            libro.active.append(list(fila.values()))
        libro.save(self.excel)
        libro.close()
        registro = RegistroConsentimientos(self.carpeta)
        registro.iniciar(datos)
        for i, fila in enumerate(datos, 1):
            ruta = self.carpeta / f'{i}.pdf'
            ruta.write_bytes(b'%PDF-1.7\n' + fila['nombre'].encode())
            registro.agregar(i, fila, ruta)
        registro.finalizar()
        self.config = {'ruta_excel': str(self.excel), 'gmail_usuario': 'cuenta@example.test',
                       'gmail_app_password': 'ficticia', 'asunto_base': 'SEPA', 'cuerpo_personalizado': 'Hola {nombre}'}
        self.servicio = EnviadorCorreos(self.carpeta)
        self.gmail = GmailLectura()
        self.imap = patch('modulos.envio_correos.EnviadorBorradores._conectar', return_value=self.gmail).start()
        self.smtp = patch('modulos.envio_correos.SMTPConRegistro').start()
        patch('modulos.envio_correos.time.sleep').start()
        self.addCleanup(patch.stopall)
        self.transporte = self.smtp.return_value.__enter__.return_value
        self.transporte.sendmail.side_effect = self.aceptar

    def aceptar(self, remitente, destinos, contenido):
        self.transporte.respuesta_data = (250, b'2.0.0 OK prueba-servidor')
        self.gmail.mensajes[str(len(self.gmail.mensajes) + 1).encode()] = contenido
        return {}

    def historial(self):
        return HistorialBorradores(self.carpeta, self.config['gmail_usuario'])

    def test_envia_adjuntos_propios_verifica_enviados_y_guarda_respuesta(self):
        plan = self.servicio.preparar_envio(self.config)
        self.imap.assert_not_called()
        self.smtp.assert_not_called()
        resultado = self.servicio.enviar(plan, self.config)
        self.assertEqual((resultado['enviados'], resultado['verificados'], resultado['pendientes']), (2, 2, 0))
        self.assertFalse(resultado['errores'] or resultado['avisos'])
        for llamada, item in zip(self.transporte.sendmail.call_args_list, plan['mensajes']):
            self.assertEqual(llamada.args, (self.config['gmail_usuario'], [item['email']], item['contenido']))
            mensaje = BytesParser(policy=policy.default).parsebytes(llamada.args[2])
            self.assertEqual(next(mensaje.iter_attachments()).get_payload(decode=True), (self.carpeta / item['archivo']).read_bytes())
        for registro in self.historial().registros.values():
            self.assertEqual(registro['smtp_codigo'], 250)
            self.assertEqual(registro['smtp_respuesta'], '2.0.0 OK prueba-servidor')
            self.assertTrue(registro['enviados_utc'])
        self.assertTrue((self.carpeta / 'informe_envios.csv').exists())
        self.assertTrue(set(self.gmail.consultas) <= {'SEARCH', 'FETCH'})
        self.smtp.reset_mock()
        segundo = self.servicio.preparar_envio(self.config)
        self.assertEqual(segundo['omitidos'], 2)
        self.servicio.enviar(segundo, self.config)
        self.smtp.assert_not_called()

    def test_identificador_nuevo_distinto_del_borrador_preexistente(self):
        originales = self.servicio._mensajes(self.config)
        h = self.historial()
        for m in originales:
            h.registrar(m)
        plan = self.servicio.preparar_envio(self.config)
        for original, item in zip(originales, plan['mensajes']):
            nuevo = BytesParser(policy=policy.default).parsebytes(item['contenido'])
            self.assertNotEqual(nuevo['Message-ID'], original['Message-ID'])
        self.assertEqual(self.servicio.enviar(plan, self.config)['verificados'], 2)

    def test_ultimo_pdf_modificado_bloquea_todo_antes_de_conectar(self):
        plan = self.servicio.preparar_envio(self.config)
        (self.carpeta / '2.pdf').write_bytes(b'alterado')
        with self.assertRaises(ValueError):
            self.servicio.enviar(plan, self.config)
        self.imap.assert_not_called()
        self.smtp.assert_not_called()

    def test_plan_con_destino_o_contenido_alterado_se_rechaza(self):
        for clave, valor in [('email', 'otro@example.test'), ('contenido', b'otro correo')]:
            with self.subTest(clave=clave):
                plan = self.servicio.preparar_envio(self.config)
                plan['mensajes'][-1][clave] = valor
                with self.assertRaises(ValueError):
                    self.servicio.enviar(plan, self.config)
        self.smtp.assert_not_called()

    def test_no_repite_envios_de_la_version_anterior(self):
        h = self.historial()
        for mensaje in self.servicio._mensajes(self.config):
            h.registrar(mensaje)
            h.marcar(next(mensaje.iter_attachments()).get_filename(), 'enviado')
        self.assertEqual(self.servicio.preparar_envio(self.config)['omitidos'], 2)
        self.smtp.assert_not_called()

    def test_timeout_persiste_incertidumbre_y_bloquea_reintentos(self):
        plan = self.servicio.preparar_envio(self.config)
        self.transporte.sendmail.side_effect = TimeoutError('Sin respuesta')
        r = self.servicio.enviar(plan, self.config)
        self.assertEqual(r['enviados'], 0)
        self.assertTrue(r['errores'])
        self.assertEqual(self.historial().registros['1.pdf']['estado'], 'sin_confirmar')
        with self.assertRaisesRegex(ValueError, 'sin confirmar'):
            self.servicio.preparar_envio(self.config)
        with self.assertRaisesRegex(ValueError, 'sin confirmar'):
            self.servicio.enviar(plan, self.config)
        self.transporte.sendmail.assert_called_once()

    def test_sin_copia_enviada_detiene_lote_sin_repetir_aceptado(self):
        self.gmail.visible = False
        r = self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.assertEqual((r['enviados'], r['verificados'], r['pendientes']), (1, 0, 1))
        self.assertTrue(r['avisos'])
        self.assertEqual(self.historial().registros['1.pdf']['estado'], 'enviado')
        self.assertNotIn('enviados_utc', self.historial().registros['1.pdf'])
        self.assertEqual(self.servicio.preparar_envio(self.config)['omitidos'], 1)
        self.transporte.sendmail.assert_called_once()

    def test_copia_que_aparece_con_retraso_no_reenvia(self):
        self.gmail.retardo = 2
        r = self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.assertEqual(r['verificados'], 2)
        self.assertEqual(self.transporte.sendmail.call_count, 2)

    def test_fallo_imap_posterior_no_pierde_aceptacion_smtp(self):
        self.gmail.fallar_busqueda = True
        r = self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.assertEqual(r['enviados'], 1)
        self.assertTrue(r['avisos'])
        self.assertEqual(self.historial().registros['1.pdf']['smtp_codigo'], 250)

    def test_copia_con_pdf_incorrecto_no_se_marca_verificada(self):
        def aceptar_alterado(remitente, destinos, contenido):
            mensaje = BytesParser(policy=policy.default).parsebytes(contenido)
            next(mensaje.iter_attachments()).set_payload('cGRmIGluY29ycmVjdG8=')
            return self.aceptar(remitente, destinos, mensaje.as_bytes())
        self.transporte.sendmail.side_effect = aceptar_alterado
        r = self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.assertEqual((r['enviados'], r['verificados']), (1, 0))
        self.assertTrue(r['avisos'])

    def test_sin_carpeta_enviados_no_envia(self):
        self.gmail.carpetas = []
        with self.assertRaisesRegex(ValueError, 'Enviados'):
            self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.smtp.assert_not_called()

    def test_fallo_al_persistir_antes_de_enviar_no_abre_envio(self):
        plan = self.servicio.preparar_envio(self.config)
        with patch.object(HistorialBorradores, 'guardar', side_effect=OSError('Disco lleno')):
            r = self.servicio.enviar(plan, self.config)
        self.transporte.sendmail.assert_not_called()
        self.assertTrue(r['errores'])

    def test_transporte_conserva_respuesta_data_sin_red(self):
        transporte = object.__new__(SMTPConRegistro)
        with patch('smtplib.SMTP.data', return_value=(250, b'2.0.0 OK identificador')):
            self.assertEqual(transporte.data(b'mensaje'), (250, b'2.0.0 OK identificador'))
        self.assertEqual(transporte.respuesta_data, (250, b'2.0.0 OK identificador'))

    def test_fallo_al_cerrar_smtp_conserva_resultados_e_informe(self):
        self.smtp.return_value.__exit__.side_effect = OSError('Fallo de cierre')
        r = self.servicio.enviar(self.servicio.preparar_envio(self.config), self.config)
        self.assertEqual((r['enviados'], r['verificados']), (2, 2))
        self.assertTrue(r['avisos'])
        self.assertTrue((self.carpeta / 'informe_envios.csv').exists())

    def test_pdf_intercambiado_aunque_se_recalcule_huella_se_rechaza(self):
        from modulos.envio_borradores import huella
        plan = self.servicio.preparar_envio(self.config)
        mensaje = BytesParser(policy=policy.default).parsebytes(plan['mensajes'][0]['contenido'])
        otro = BytesParser(policy=policy.default).parsebytes(plan['mensajes'][1]['contenido'])
        mensaje.set_payload([mensaje.get_body(), next(otro.iter_attachments())])
        plan['mensajes'][0]['contenido'] = mensaje.as_bytes()
        plan['mensajes'][0]['huella'] = huella(mensaje.as_bytes())
        with self.assertRaisesRegex(ValueError, 'PDF'):
            self.servicio.enviar(plan, self.config)
        self.smtp.assert_not_called()


if __name__ == '__main__':
    unittest.main()
