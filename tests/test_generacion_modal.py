"""Regresión del bloqueo de Tk durante una conversión lenta."""

import tempfile
import threading
import tkinter as tk
import unittest
from pathlib import Path
from time import monotonic
from unittest.mock import patch

from openpyxl import Workbook

from app_consentimientos import AplicacionConsentimientos
from modulos.registro_consentimientos import RegistroConsentimientos


class GeneracionModalTest(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk no está disponible: {exc}")
        self.root.withdraw()
        self.temporal = tempfile.TemporaryDirectory()
        self.carpeta = Path(self.temporal.name)
        with patch.object(AplicacionConsentimientos, "_cargar_config", return_value={}), \
                patch.object(self.root, "state"):
            self.app = AplicacionConsentimientos(self.root)
        libro = Workbook()
        libro.active.append(["nombre", "email"])
        libro.active.append(["Ana", "ana@example.test"])
        libro.active.append(["Luis", "luis@example.test"])
        excel = self.carpeta / "datos.xlsx"
        libro.save(excel)
        libro.close()
        plantilla = self.carpeta / "plantilla.docx"
        plantilla.write_bytes(b"plantilla simulada")
        self.app.archivo_excel.set(str(excel))
        self.app.plantilla_word.set(str(plantilla))
        self.app.carpeta_salida.set(str(self.carpeta))
        self.app.mapeo_campos = {"nombre_deudor": "nombre"}
        self.liberar = [threading.Event(), threading.Event()]
        self.iniciados = [threading.Event(), threading.Event()]
        self.patches = [
            patch("subprocess.run"), patch("time.sleep"),
            patch("app_consentimientos.messagebox"),
            patch("modulos.generador_word.GeneradorConsentimientos"),
        ]
        self.mocks = [p.start() for p in self.patches]
        self.avisos = self.mocks[2]
        self.fabrica = self.mocks[3]
        self.hilos_conversion = []

        def convertir(datos):
            i = len(self.hilos_conversion)
            self.hilos_conversion.append(threading.get_ident())
            self.iniciados[i].set()
            if not self.liberar[i].wait(timeout=5):
                raise TimeoutError("La prueba no liberó la conversión")
            ruta = self.carpeta / f"pdf_{i}.pdf"
            ruta.write_bytes(datos["nombre_deudor"].encode())
            return str(ruta)

        self.fabrica.return_value.generar_consentimiento.side_effect = convertir

    def tearDown(self):
        for evento in self.liberar:
            evento.set()
        hilo = getattr(self.app, "_hilo_generacion", None)
        if hilo:
            hilo.join(timeout=6)
        hilo_borradores = getattr(self.app, "_hilo_borradores", None)
        if hilo_borradores:
            hilo_borradores.join(timeout=6)
        hilo_envio = getattr(self.app, "_hilo_envio", None)
        if hilo_envio:
            hilo_envio.join(timeout=6)
        for p in reversed(self.patches):
            p.stop()
        # Retirar callbacks antes de destruir Tk para que no afecten a otra prueba.
        for callback in self.root.tk.call("after", "info"):
            self.root.after_cancel(callback)
        self.root.destroy()
        self.temporal.cleanup()

    def esperar(self, condicion, timeout=3):
        limite = monotonic() + timeout
        pausa = threading.Event()
        while not condicion() and monotonic() < limite:
            self.root.update()
            pausa.wait(0.01)
        self.assertTrue(condicion(), "No llegó el estado esperado de la generación")

    def test_modal_responde_y_avanza_mientras_word_sigue_trabajando(self):
        self.assertIsNone(self.app.ventana_progreso)
        self.assertFalse(hasattr(self.app, "barra_progreso"))
        self.app.generar_consentimientos()
        modal = self.app.ventana_progreso
        self.assertTrue(modal.winfo_exists())
        self.assertEqual(self.root.grab_current(), modal)
        self.esperar(self.iniciados[0].is_set)
        self.assertNotEqual(self.hilos_conversion[0], threading.get_ident())
        # Un segundo clic no puede lanzar otro proceso ni limpiar los PDFs en curso.
        self.app.generar_consentimientos()
        self.assertEqual(self.fabrica.call_count, 1)
        self.app._cerrar_aplicacion()
        self.assertTrue(self.root.winfo_exists())

        latidos = []
        def latido():
            latidos.append(monotonic())
            if self.app._generando:
                self.root.after(20, latido)
        self.root.after(20, latido)
        self.esperar(lambda: len(latidos) >= 4)
        self.esperar(lambda: self.app.progreso_tiempo.get().startswith("Transcurrido: 2 s"))
        self.assertEqual(float(self.app.barra_progreso["value"]), 0)

        self.liberar[0].set()
        self.esperar(self.iniciados[1].is_set)
        self.esperar(lambda: float(self.app.barra_progreso["value"]) == 50)
        self.assertIn("1 pendientes", self.app.progreso_detalle.get())
        self.assertIn("1 PDFs generados", self.app.progreso_detalle.get())
        self.assertTrue(self.app._hilo_generacion.is_alive())
        self.liberar[1].set()
        self.esperar(lambda: not self.app._generando)
        self.assertIsNone(self.app.ventana_progreso)
        self.assertFalse(modal.winfo_exists())
        self.assertIsNone(self.root.grab_current())
        self.assertEqual(str(self.app.btn_generar["state"]), "normal")
        self.avisos.showinfo.assert_called_once()
        self.avisos.showerror.assert_not_called()
        self.assertEqual(len(list(self.carpeta.glob("*.pdf"))), 2)

    def test_error_fatal_cierra_modal_y_permite_reintentar(self):
        self.fabrica.side_effect = RuntimeError("Word no disponible")
        self.app.generar_consentimientos()
        modal = self.app.ventana_progreso
        self.esperar(lambda: not self.app._generando)
        self.assertFalse(modal.winfo_exists())
        self.assertEqual(str(self.app.btn_generar["state"]), "normal")
        self.avisos.showerror.assert_called_once()
        registro = RegistroConsentimientos(self.carpeta)
        with self.assertRaisesRegex(ValueError, "Borradores bloqueados"):
            registro.validar([])
        self.fabrica.side_effect = None
        for evento in self.liberar:
            evento.set()
        self.app.generar_consentimientos()
        self.esperar(lambda: not self.app._generando)
        self.avisos.showinfo.assert_called_once()

    def test_fallo_de_un_pdf_se_muestra_y_bloquea_envio(self):
        self.fabrica.return_value.generar_consentimiento.side_effect = RuntimeError("Error de conversión")
        self.app.generar_consentimientos()
        self.esperar(lambda: not self.app._generando)
        self.assertIn("2 errores", self.app.progreso_detalle.get())
        self.assertIn("0 PDFs generados", self.app.progreso_detalle.get())
        self.assertIsNone(self.app.ventana_progreso)
        with self.assertRaisesRegex(ValueError, "Borradores bloqueados"):
            RegistroConsentimientos(self.carpeta).validar([])

    def test_crear_borradores_se_ejecuta_en_segundo_plano_y_cierra_modal(self):
        self.app.gmail_usuario.set("prueba@example.test")
        self.app.gmail_app_password.set("clave-ficticia")
        self.assertFalse(hasattr(self.app, "enviar_emails_gmail"))
        hilos = []

        def crear(**configuracion):
            hilos.append(threading.get_ident())
            configuracion["progreso"](1, 2, 1, 0)
            self.iniciados[0].set()
            self.liberar[0].wait(timeout=5)
            return {"total": 2, "creados": 2, "existentes": 0, "fallidos": [], "pendientes": 0}

        with patch.object(self.app, "_guardar_config"), \
                patch("modulos.preparador_emails.PreparadorEmails.crear_borradores_gmail", side_effect=crear) as servicio:
            self.app.crear_borradores_gmail()
            modal = self.app.ventana_progreso
            self.assertEqual(modal.title(), "Creando borradores en Gmail")
            self.esperar(self.iniciados[0].is_set)
            self.esperar(lambda: float(self.app.barra_progreso["value"]) == 50)
            self.assertNotEqual(hilos[0], threading.get_ident())
            self.app.crear_borradores_gmail()
            servicio.assert_called_once()
            self.liberar[0].set()
            self.esperar(lambda: not self.app._generando)
        self.assertFalse(modal.winfo_exists())
        self.assertEqual(str(self.app.btn_borradores["state"]), "normal")
        self.assertIn("El programa no ha enviado", self.avisos.showinfo.call_args.args[1])

    def test_error_de_gmail_libera_la_modal(self):
        self.app.gmail_usuario.set("prueba@example.test")
        self.app.gmail_app_password.set("clave-ficticia")
        with patch.object(self.app, "_guardar_config"), \
                patch("modulos.preparador_emails.PreparadorEmails.crear_borradores_gmail",
                      side_effect=RuntimeError("No se pudo conectar con Gmail")):
            self.app.crear_borradores_gmail()
            self.esperar(lambda: not self.app._generando)
        self.assertIsNone(self.app.ventana_progreso)
        self.assertEqual(str(self.app.btn_borradores["state"]), "normal")
        self.avisos.showerror.assert_called_once()

    def test_envio_requiere_confirmar_la_lista_de_borradores(self):
        self.app.gmail_usuario.set("prueba@example.test")
        self.app.gmail_app_password.set("clave-ficticia")
        plan = {"cuenta": "prueba@example.test", "omitidos": 0,
                "mensajes": [{"email": "ana@example.test", "asunto": "Texto de Gmail", "archivo": "ana.pdf"}]}
        resultado = {"enviados": 1, "omitidos": 0, "pendientes": 0, "errores": [], "avisos": []}

        def botones(widget):
            encontrados = []
            for hijo in widget.winfo_children():
                if hijo.winfo_class() == "TButton":
                    encontrados.append(hijo)
                encontrados.extend(botones(hijo))
            return encontrados

        def confirmacion_visible():
            ventana = self.app.ventana_progreso
            return ventana is not None and ventana.title() == "Confirmar envío de borradores revisados"

        with patch("modulos.envio_borradores.EnviadorBorradores.preparar_envio", return_value=plan), \
                patch("modulos.envio_borradores.EnviadorBorradores.enviar", return_value=resultado) as enviar:
            self.app.enviar_borradores_gmail()
            self.esperar(confirmacion_visible)
            enviar.assert_not_called()
            cancelar = next(b for b in botones(self.app.ventana_progreso) if b["text"] == "Cancelar")
            cancelar.invoke()
            self.assertFalse(self.app._generando)
            enviar.assert_not_called()
            self.app.enviar_borradores_gmail()
            self.esperar(confirmacion_visible)
            enviar.assert_not_called()
            confirmar = next(b for b in botones(self.app.ventana_progreso) if str(b["text"]).startswith("Enviar los"))
            confirmar.invoke()
            self.esperar(lambda: not self.app._generando)
            enviar.assert_called_once()
            self.assertIs(enviar.call_args.args[0], plan)
        self.assertIsNone(self.app.ventana_progreso)
        self.assertIn("Enviados: 1", self.avisos.showinfo.call_args.args[1])

    def test_fallo_en_verificacion_no_muestra_confirmacion_ni_envia(self):
        self.app.gmail_usuario.set("prueba@example.test")
        self.app.gmail_app_password.set("clave-ficticia")
        with patch("modulos.envio_borradores.EnviadorBorradores.preparar_envio", side_effect=ValueError("PDF incorrecto")), \
                patch("modulos.envio_borradores.EnviadorBorradores.enviar") as enviar:
            self.app.enviar_borradores_gmail()
            self.esperar(lambda: not self.app._generando)
            enviar.assert_not_called()
        self.assertIsNone(self.app.ventana_progreso)
        self.avisos.showerror.assert_called_once()


if __name__ == "__main__":
    unittest.main()
