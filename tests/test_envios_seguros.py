import csv
import tempfile
import unittest
from pathlib import Path
from email import policy
from email.parser import BytesParser
from unittest.mock import Mock, patch

from openpyxl import Workbook

from modulos.generador_word import GeneradorConsentimientos
from modulos.preparador_emails import PreparadorEmails
from modulos.registro_consentimientos import RegistroConsentimientos


class EnviosSegurosTest(unittest.TestCase):
    def setUp(self):
        self.temporal = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporal.cleanup)
        self.carpeta = Path(self.temporal.name)
        self.excel = self.carpeta / "datos.xlsx"
        self.datos = [
            {"nombre": "Ana", "email": "ana@example.test", "curso": "1"},
            {"nombre": "Anabel", "email": "anabel@example.test", "curso": "1"},
            {"nombre": "Ana", "email": "otra@example.test", "curso": "1"},
        ]
        self.guardar_excel()
        self.registro = RegistroConsentimientos(self.carpeta)
        self.preparador = PreparadorEmails(str(self.carpeta))
        self.imap = patch("modulos.preparador_emails.imaplib.IMAP4_SSL").start()
        self.smtp = patch("smtplib.SMTP_SSL", side_effect=AssertionError("SMTP prohibido")).start()
        self.smtp_plain = patch("smtplib.SMTP", side_effect=AssertionError("SMTP prohibido")).start()
        self.addCleanup(self.smtp.assert_not_called)
        self.addCleanup(self.smtp_plain.assert_not_called)
        self.addCleanup(patch.stopall)
        self.servidor = self.imap.return_value.__enter__.return_value
        self.servidor.list.return_value = ("OK", [b'(\\HasNoChildren \\Drafts) "/" "[Gmail]/Borradores"'])
        self.servidor.select.return_value = ("OK", [b"0"])
        self.servidor.uid.return_value = ("OK", [b""])
        self.servidor.append.return_value = ("OK", [b"guardado"])

    def guardar_excel(self):
        libro = Workbook()
        hoja = libro.active
        hoja.append(list(self.datos[0]))
        for registro in self.datos:
            hoja.append(list(registro.values()))
        libro.save(self.excel)
        libro.close()

    def generar_lote(self, completo=True):
        self.registro.iniciar(self.datos)
        self.pdfs = []
        for numero, registro in enumerate(self.datos, 1):
            pdf = self.carpeta / f"Ana_curso1_{numero}.pdf"
            pdf.write_bytes(b"%PDF-1.7\n" + registro["email"].encode())
            self.registro.agregar(numero, registro, pdf)
            self.pdfs.append(pdf)
        if completo:
            self.registro.finalizar()

    def crear_gmail(self):
        return self.preparador.crear_borradores_gmail(
            str(self.excel), "remitente@example.test", "clave-ficticia"
        )

    def comprobar_bloqueo(self):
        with self.assertRaisesRegex(ValueError, "Borradores bloqueados"):
            self.crear_gmail()
        self.imap.assert_not_called()

    def test_adjuntos_exactos_con_nombres_y_valores_compartidos(self):
        self.generar_lote()
        (self.carpeta / "Ana_antiguo.pdf").write_bytes(b"PDF antiguo ajeno al lote")
        resultado = self.crear_gmail()
        self.assertEqual(resultado["creados"], 3)
        for llamada, registro, pdf in zip(self.servidor.append.call_args_list, self.datos, self.pdfs):
            self.assertEqual(llamada.args[0], b'"[Gmail]/Borradores"')
            self.assertEqual(llamada.args[1], r"(\Draft)")
            mensaje = BytesParser(policy=policy.default).parsebytes(llamada.args[3])
            self.assertEqual(str(mensaje["To"]), registro["email"])
            adjuntos = list(mensaje.iter_attachments())
            self.assertEqual(len(adjuntos), 1)
            self.assertEqual(adjuntos[0].get_payload(decode=True), pdf.read_bytes())
            self.assertEqual(adjuntos[0].get_filename(), pdf.name)

    def test_borradores_y_csv_usan_la_misma_asociacion(self):
        self.generar_lote()
        borradores = self.preparador.crear_borradores_email(str(self.excel))
        self.assertEqual([b["archivo"] for b in borradores], [str(p) for p in self.pdfs])
        with (self.carpeta / "emails_lista.csv").open(encoding="utf-8", newline="") as archivo:
            filas = list(csv.DictReader(archivo))
        self.assertEqual([f["Archivo Adjunto"] for f in filas], [p.name for p in self.pdfs])
        self.imap.assert_not_called()

    def test_pdf_unico_sin_asociacion_no_se_reutiliza(self):
        (self.carpeta / "Ana.pdf").write_bytes(b"PDF de Ana")
        self.comprobar_bloqueo()
        with self.assertRaises(ValueError):
            self.preparador.crear_borradores_email(str(self.excel))

    def test_generacion_incompleta(self):
        self.generar_lote(completo=False)
        self.comprobar_bloqueo()

    def test_registro_fallido_no_desplaza_los_adjuntos(self):
        self.generar_lote()
        self.registro.datos["archivos"].pop(1)
        self.registro.finalizar()
        self.comprobar_bloqueo()

    def test_excel_modificado(self):
        self.generar_lote()
        self.datos[1]["email"] = "cambiado@example.test"
        self.guardar_excel()
        self.comprobar_bloqueo()

    def test_excel_reordenado(self):
        self.generar_lote()
        self.datos.reverse()
        self.guardar_excel()
        self.comprobar_bloqueo()

    def test_ultimo_pdf_ausente_bloquea_todo(self):
        self.generar_lote()
        self.pdfs[-1].unlink()
        self.comprobar_bloqueo()

    def test_pdf_modificado(self):
        self.generar_lote()
        self.pdfs[-1].write_bytes(self.pdfs[0].read_bytes())
        self.comprobar_bloqueo()

    def test_asociaciones_al_mismo_archivo(self):
        self.generar_lote()
        entradas = self.registro.datos["archivos"]
        entradas[-1]["archivo"] = entradas[0]["archivo"]
        entradas[-1]["sha256"] = entradas[0]["sha256"]
        self.registro._guardar()
        self.comprobar_bloqueo()

    def test_contenido_duplicado_con_nombres_distintos(self):
        self.generar_lote()
        self.pdfs[-1].write_bytes(self.pdfs[0].read_bytes())
        entradas = self.registro.datos["archivos"]
        entradas[-1]["sha256"] = entradas[0]["sha256"]
        self.registro._guardar()
        self.comprobar_bloqueo()

    def test_asociaciones_intercambiadas(self):
        self.generar_lote()
        self.registro.datos["archivos"].reverse()
        self.registro._guardar()
        self.comprobar_bloqueo()

    def test_registro_corrupto(self):
        self.registro.ruta.write_text("{", encoding="utf-8")
        self.comprobar_bloqueo()

    def test_destinatarios_invalidos_bloquean_antes_de_imap(self):
        for email in ["", "sin-arroba", "a@example.test,b@example.test", "a@example.test\nBcc: b@example.test"]:
            with self.subTest(email=email):
                self.datos[-1]["email"] = email
                self.guardar_excel()
                self.generar_lote()
                self.comprobar_bloqueo()

    def test_columnas_correo_ambiguas(self):
        for registro in self.datos:
            registro["correo alternativo"] = "alternativo@example.test"
        self.guardar_excel()
        self.generar_lote()
        self.comprobar_bloqueo()

    def test_borrador_conserva_los_bytes_validados_aunque_el_archivo_cambie(self):
        self.generar_lote()
        original = self.pdfs[0].read_bytes()
        self.servidor.login.side_effect = lambda *args: self.pdfs[0].write_bytes(b"cambio posterior")
        self.crear_gmail()
        mensaje = BytesParser(policy=policy.default).parsebytes(self.servidor.append.call_args_list[0].args[3])
        self.assertEqual(next(mensaje.iter_attachments()).get_payload(decode=True), original)

    def test_nombres_unicos_incluso_para_homonimos(self):
        generador = GeneradorConsentimientos.__new__(GeneradorConsentimientos)
        nombres = {generador._generar_nombre_archivo(self.datos[0]) for _ in range(100)}
        self.assertEqual(len(nombres), 100)

    def test_borradores_existentes_no_se_duplican(self):
        self.generar_lote()
        self.servidor.uid.side_effect = [("OK", [b"42"]), ("OK", [b""]), ("OK", [b"43"])]
        resultado = self.crear_gmail()
        self.assertEqual(resultado["creados"], 1)
        self.assertEqual(resultado["existentes"], 2)
        self.assertEqual(resultado["pendientes"], 0)
        self.servidor.append.assert_called_once()
        mensaje = BytesParser(policy=policy.default).parsebytes(self.servidor.append.call_args.args[3])
        self.assertEqual(str(mensaje["To"]), self.datos[1]["email"])

    def test_reintento_busca_el_mismo_identificador(self):
        self.generar_lote()
        self.crear_gmail()
        primera_busqueda = self.servidor.uid.call_args_list[:]
        self.servidor.uid.reset_mock()
        self.servidor.uid.return_value = ("OK", [b"42"])
        self.crear_gmail()
        self.assertEqual(primera_busqueda, self.servidor.uid.call_args_list)
        self.assertEqual(self.servidor.append.call_count, 3)

    def test_carpeta_borradores_se_detecta_sin_depender_del_idioma(self):
        self.generar_lote()
        self.servidor.list.return_value = ("OK", [
            b'(\\HasNoChildren) "/" "INBOX"',
            b'(\\HasNoChildren \\Sent) "/" "[Gmail]/Sent Mail"',
            b'(\\HasNoChildren \\Drafts) "/" "[Google Mail]/Entw&APw-rfe"',
        ])
        self.crear_gmail()
        self.servidor.select.assert_called_once_with(b'"[Google Mail]/Entw&APw-rfe"', readonly=True)
        self.assertEqual(self.servidor.append.call_args.args[0], b'"[Google Mail]/Entw&APw-rfe"')

    def test_sin_carpeta_borradores_no_escribe_en_otra(self):
        self.generar_lote()
        self.servidor.list.return_value = ("OK", [b'(\\Inbox) "/" "INBOX"'])
        with self.assertRaisesRegex(ValueError, "carpeta Borradores"):
            self.crear_gmail()
        self.servidor.append.assert_not_called()

    def test_rechazo_de_gmail_detiene_el_lote(self):
        self.generar_lote()
        self.servidor.append.side_effect = [("OK", [b"guardado"]), ("NO", [b"quota"])]
        resultado = self.crear_gmail()
        self.assertEqual(resultado["creados"], 1)
        self.assertEqual(len(resultado["fallidos"]), 1)
        self.assertEqual(resultado["fallidos"][0]["numero"], 2)
        self.assertEqual(resultado["pendientes"], 1)
        self.assertEqual(self.servidor.append.call_count, 2)

    def test_timeout_no_reintenta_ni_confirma_creacion(self):
        self.generar_lote()
        self.servidor.append.side_effect = TimeoutError("Tiempo agotado")
        resultado = self.crear_gmail()
        self.assertEqual(resultado["creados"], 0)
        self.assertEqual(len(resultado["fallidos"]), 1)
        self.assertEqual(resultado["pendientes"], 2)
        self.servidor.append.assert_called_once()

    def test_error_de_busqueda_no_crea_borradores_a_ciegas(self):
        self.generar_lote()
        self.servidor.uid.return_value = ("NO", [b"error"])
        resultado = self.crear_gmail()
        self.assertEqual(len(resultado["fallidos"]), 1)
        self.servidor.append.assert_not_called()

    def test_no_existe_el_metodo_de_envio_directo(self):
        self.assertFalse(hasattr(self.preparador, "enviar_emails_gmail"))

    def test_nueva_generacion_limpia_pdfs_y_borradores_sin_borrar_entradas(self):
        self.generar_lote()
        self.preparador.crear_borradores_email(str(self.excel))
        mayusculas = self.carpeta / "anterior.PDF"
        mayusculas.write_bytes(b"anterior")
        plantilla = self.carpeta / "plantilla.docx"
        plantilla.write_bytes(b"plantilla")
        notas = self.carpeta / "notas.txt"
        notas.write_text("conservar", encoding="utf-8")
        subcarpeta = self.carpeta / "archivo"
        subcarpeta.mkdir()
        archivado = subcarpeta / "guardado.pdf"
        archivado.write_bytes(b"conservar")
        eliminados = self.registro.iniciar(self.datos, limpiar_anteriores=True)
        self.assertEqual(eliminados, 6)
        self.assertFalse(mayusculas.exists())
        self.assertFalse(any(pdf.exists() for pdf in self.pdfs))
        self.assertFalse((self.carpeta / "emails_lista.csv").exists())
        self.assertFalse((self.carpeta / "emails_para_enviar.txt").exists())
        self.assertTrue(all(p.exists() for p in (self.excel, plantilla, notas, archivado)))
        self.comprobar_bloqueo()

    def test_pdf_bloqueado_detiene_la_limpieza_e_invalida_el_lote(self):
        self.generar_lote()
        with patch.object(Path, "unlink", side_effect=PermissionError("archivo abierto")):
            with self.assertRaisesRegex(ValueError, "Cierra el archivo"):
                self.registro.iniciar(self.datos, limpiar_anteriores=True)
        self.comprobar_bloqueo()

    def test_generacion_app_asocia_filas_originales_con_datos_mapeados(self):
        from app_consentimientos import AplicacionConsentimientos

        app = AplicacionConsentimientos.__new__(AplicacionConsentimientos)
        app.archivo_excel = Mock(get=Mock(return_value=str(self.excel)))
        app.plantilla_word = Mock(get=Mock(return_value="plantilla.docx"))
        app.carpeta_salida = Mock(get=Mock(return_value=str(self.carpeta)))
        app.mapeo_campos = {"nombre_deudor": "nombre"}
        app.valores_fijos = {}
        app.validar_archivos = Mock(return_value=True)
        app.agregar_log = Mock()
        app.root = Mock()
        app.btn_generar = Mock()
        app.barra_progreso = Mock()
        app.progreso_estado = Mock()
        app.progreso_detalle = Mock()
        app.progreso_tiempo = Mock()
        app._crear_modal_progreso_pdf = Mock()
        app.ventana_progreso = Mock()
        anterior = self.carpeta / "lote_anterior.pdf"
        anterior.write_bytes(b"eliminar antes de generar")
        originales = []

        def generar(registro):
            self.assertFalse(anterior.exists())
            originales.append(registro)
            pdf = self.carpeta / f"nuevo_{len(originales)}.pdf"
            pdf.write_bytes(f"PDF {len(originales)} {registro['nombre_deudor']}".encode())
            return str(pdf)

        with patch("modulos.generador_word.GeneradorConsentimientos") as fabrica, \
                patch("subprocess.run"), patch("time.sleep"), \
                patch("app_consentimientos.messagebox") as avisos:
            fabrica.return_value.generar_consentimiento.side_effect = generar
            app.generar_consentimientos()
            app._hilo_generacion.join(timeout=5)
            self.assertFalse(app._hilo_generacion.is_alive())
            app._procesar_eventos_generacion()
        avisos.showerror.assert_not_called()
        self.assertEqual([r["nombre_deudor"] for r in originales], [r["nombre"] for r in self.datos])
        self.assertEqual(len(self.registro.validar(self.datos)), 3)
        self.assertFalse(anterior.exists())
        self.assertTrue(self.excel.exists())
        self.assertEqual(self.crear_gmail()["creados"], 3)
        app.barra_progreso.configure.assert_called_with(value=100)
        app.progreso_detalle.set.assert_called_with(
            "3 de 3 procesados (100 %) · 3 PDFs generados · 0 pendientes · 0 errores"
        )
        app.btn_generar.configure.assert_called_with(state="normal")


if __name__ == "__main__":
    unittest.main()
