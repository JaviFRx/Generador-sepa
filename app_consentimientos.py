"""
Sistema de Generación de Consentimientos
Aplicación para leer datos de Excel, generar consentimientos desde plantilla Word
y prepararlos para envío por correo electrónico.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import shutil
import re
import webbrowser
from pathlib import Path
from datetime import datetime
import urllib.request
import json
import urllib.parse
import base64
import hashlib
from time import monotonic
from math import ceil
from queue import Queue, Empty
from threading import Thread
from cryptography.fernet import Fernet


# Pais que se usa cuando el Excel no aporta uno valido
PAIS_POR_DEFECTO = "España"

# Cabeceras tipicas de columnas de marca temporal (Google Forms, exportaciones)
COLUMNAS_MARCA_TEMPORAL = (
    'marca temporal', 'marca de temps', 'timestamp',
    'fecha y hora', 'data i hora', 'hora de envio', "hora d'enviament",
)


# Guia de uso publicada (se abre desde el boton "Como funciona")
URL_DOCUMENTACION = "https://claude.ai/code/artifact/55f3fb72-91e1-4b64-aff2-8b81cbb94cb8"


def ruta_base_app() -> Path:
    """Carpeta donde vive la aplicacion: junto al .exe o junto al .py"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def ruta_recursos() -> Path:
    """Carpeta de recursos empaquetados (_MEIPASS dentro del .exe)"""
    if getattr(sys, 'frozen', False):
        return Path(getattr(sys, '_MEIPASS', Path(sys.executable).parent))
    return Path(__file__).parent


def buscar_plantilla_default() -> Path:
    """Plantilla del programa en una ruta estable, nunca dentro de la carpeta temporal del .exe"""
    junto_al_exe = ruta_base_app() / "docs" / "plantilla_domiciliacion_sepa.docx"
    if junto_al_exe.exists():
        return junto_al_exe

    empaquetada = ruta_recursos() / "docs" / "plantilla_domiciliacion_sepa.docx"
    if empaquetada.exists():
        # _MEIxxxx se borra al cerrar el programa: copiarla a una carpeta permanente
        destinos = [
            junto_al_exe,
            Path.home() / "SEPAS" / "docs" / "plantilla_domiciliacion_sepa.docx",
        ]
        for destino in destinos:
            try:
                destino.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(empaquetada, destino)
                return destino
            except OSError:
                continue
        return empaquetada

    return junto_al_exe


class AplicacionConsentimientos:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Consentimientos - SEPAS")
        self.root.geometry("950x750")
        self.root.state('zoomed')  # Abrir maximizado en Windows
        self.root.resizable(True, True)
        
        # Color de fondo
        self.root.configure(bg='#f5f5f5')
        
        # Archivo de configuración
        self.ruta_config = Path.home() / ".sepas_config.json"
        self.config = self._cargar_config()
        
        # Variables
        self.archivo_excel = tk.StringVar()
        # Plantilla: la guardada en configuracion si sigue existiendo, si no la del programa
        plantilla_guardada = self.config.get("plantilla_word", "")
        if "_MEI" in plantilla_guardada:
            plantilla_guardada = ""
        if plantilla_guardada and Path(plantilla_guardada).exists():
            ruta_plantilla_default = Path(plantilla_guardada)
        else:
            ruta_plantilla_default = buscar_plantilla_default()
        self.plantilla_word = tk.StringVar(value=str(ruta_plantilla_default))
        carpeta_guardada = self.config.get("carpeta_salida", "")
        self.carpeta_salida = tk.StringVar(
            value=carpeta_guardada or str(Path.cwd() / "consentimientos_generados"))
        self.gmail_usuario = tk.StringVar(value=self.config.get("gmail_usuario", ""))
        self.gmail_app_password = tk.StringVar(value=self.config.get("gmail_app_password", ""))
        self.gmail_nombre_remitente = tk.StringVar(value=self.config.get("gmail_nombre_remitente", ""))
        self.email_asunto = tk.StringVar(value=self.config.get("email_asunto", "Mandato SEPA: Confirmación de datos para cobros periódicos"))
        self.email_cuerpo = self.config.get("email_cuerpo", "")
        
        # Mapeo de columnas Excel a campos de la plantilla
        self.mapeo_campos = {}
        self.valores_fijos = {}  # Valores fijos para campos no mapeados
        self._generando = False
        self.ventana_progreso = None
        self.root.protocol("WM_DELETE_WINDOW", self._cerrar_aplicacion)
        
        # Crear interfaz
        self.crear_interfaz()
        
    def crear_interfaz(self):
        """Crea la interfaz gráfica de la aplicación"""

        # Frame principal con scroll
        canvas = tk.Canvas(self.root, bg='#f5f5f5', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#f5f5f5', padx=20, pady=10)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Ajustar ancho dinámicamente
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", on_canvas_configure)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Bind scroll del ratón
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        main_frame = scrollable_frame
        
        # Configurar fuentes más grandes
        fuente_titulo = ('Arial', 24, 'bold')
        fuente_seccion = ('Arial', 13, 'bold')
        fuente_normal = ('Arial', 11)
        fuente_boton = ('Arial', 14, 'bold')
        fuente_boton_grande = ('Arial', 16, 'bold')
        
        # Banner de título con color de fondo
        banner_frame = tk.Frame(main_frame, bg='#673AB7', relief=tk.RAISED, bd=3)
        banner_frame.pack(fill=tk.X, pady=(0, 6))
        
        titulo = tk.Label(banner_frame, text="⚕ GENERADOR DE CONSENTIMIENTOS - SEPAS", 
                          font=fuente_titulo, bg='#673AB7', fg='white', pady=15)
        titulo.pack(fill=tk.X)
        
        # Enlace discreto a la guia de uso
        ayuda_frame = tk.Frame(main_frame, bg='#f5f5f5')
        ayuda_frame.pack(fill=tk.X, pady=(0, 14))
        
        btn_ayuda = tk.Button(ayuda_frame, text="❔ Cómo funciona",
                              command=self.abrir_documentacion,
                              font=('Arial', 9, 'underline'),
                              bg='#f5f5f5', fg='#673AB7',
                              activebackground='#f5f5f5', activeforeground='#512DA8',
                              relief=tk.FLAT, bd=0, cursor='hand2',
                              padx=4, pady=0)
        btn_ayuda.pack(side=tk.RIGHT)
        
        # Sección 1: Archivo Excel
        tk.Label(main_frame, text="➊ Archivo Excel con datos:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(10, 5))
        
        excel_frame = tk.Frame(main_frame, bg='#f5f5f5')
        excel_frame.pack(fill=tk.X, pady=(0, 20))
        
        entry_excel = ttk.Entry(excel_frame, textvariable=self.archivo_excel, 
                 width=60, font=fuente_normal)
        entry_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        btn_excel = tk.Button(excel_frame, text="📁 Seleccionar Excel", 
                  command=self.seleccionar_excel, font=fuente_boton,
                  bg='#4CAF50', fg='white', activebackground='#45a049',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_excel.pack(side=tk.LEFT)
        
        # Sección 2: Plantilla Word
        tk.Label(main_frame, text="➋ Plantilla Word (consentimiento):", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(20, 5))
        
        word_frame = tk.Frame(main_frame, bg='#f5f5f5')
        word_frame.pack(fill=tk.X, pady=(0, 20))
        
        entry_word = ttk.Entry(word_frame, textvariable=self.plantilla_word, 
                 width=60, font=fuente_normal)
        entry_word.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        btn_word = tk.Button(word_frame, text="📄 Seleccionar Plantilla", 
                  command=self.seleccionar_plantilla, font=fuente_boton,
                  bg='#2196F3', fg='white', activebackground='#0b7dda',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_word.pack(side=tk.LEFT)
        
        # Sección 3: Carpeta de salida
        tk.Label(main_frame, text="➌ Carpeta para guardar consentimientos:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(20, 5))
        
        carpeta_frame = tk.Frame(main_frame, bg='#f5f5f5')
        carpeta_frame.pack(fill=tk.X, pady=(0, 20))
        
        entry_salida = ttk.Entry(carpeta_frame, textvariable=self.carpeta_salida, 
                 width=60, font=fuente_normal)
        entry_salida.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        btn_carpeta = tk.Button(carpeta_frame, text="📂 Seleccionar Carpeta", 
                  command=self.seleccionar_carpeta, font=fuente_boton,
                  bg='#FF9800', fg='white', activebackground='#e68900',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_carpeta.pack(side=tk.LEFT)

        # Sección 4: Cuenta Gmail para envío
        tk.Label(main_frame, text="➍ Cuenta Gmail para crear borradores:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(20, 5))

        gmail_frame = tk.Frame(main_frame, bg='#f5f5f5')
        gmail_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Label(gmail_frame, text="Correo Gmail:", font=fuente_normal, bg='#f5f5f5').pack(anchor=tk.W, pady=(0, 3))
        entry_gmail = ttk.Entry(gmail_frame, textvariable=self.gmail_usuario, width=60, font=fuente_normal)
        entry_gmail.pack(fill=tk.X, pady=(0, 10))

        tk.Label(gmail_frame, text="Contraseña de aplicación:", font=fuente_normal, bg='#f5f5f5').pack(anchor=tk.W, pady=(0, 3))
        entry_gmail_pass = ttk.Entry(
            gmail_frame, textvariable=self.gmail_app_password, width=60, font=fuente_normal, show="*"
        )
        entry_gmail_pass.pack(fill=tk.X, pady=(0, 10))

        tk.Label(gmail_frame, text="Nombre remitente (opcional):", font=fuente_normal, bg='#f5f5f5').pack(anchor=tk.W, pady=(0, 3))
        entry_gmail_nombre = ttk.Entry(
            gmail_frame, textvariable=self.gmail_nombre_remitente, width=60, font=fuente_normal
        )
        entry_gmail_nombre.pack(fill=tk.X, pady=(0, 5))

        tk.Label(
            gmail_frame,
            text="Crea los borradores, revísalos en Gmail y vuelve aquí para enviarlos cuando estén correctos.",
            font=('Arial', 9),
            bg='#f5f5f5',
            fg='#666'
        ).pack(anchor=tk.W, pady=(2, 6))
        
        # Botón guardar configuración Gmail
        btn_guardar_gmail = tk.Button(
            gmail_frame,
            text="💾 Guardar Configuración",
            command=self._guardar_config,
            font=('Arial', 9),
            bg='#673AB7', fg='white',
            activebackground='#512DA8',
            cursor='hand2', relief=tk.RAISED, bd=2,
            padx=15, pady=5
        )
        btn_guardar_gmail.pack(anchor=tk.W, pady=(5, 0))
        
        # Botón principal
        btn_frame = tk.Frame(main_frame, bg='#f5f5f5')
        btn_frame.pack(pady=30)
        
        # Botón principal de generar - MÁS GRANDE Y LLAMATIVO
        self.btn_generar = tk.Button(btn_frame, 
                                     text="🗎 GENERAR\nCONSENTIMIENTOS\nEN PDF", 
                                     command=self.generar_consentimientos,
                                     font=fuente_boton,
                                     bg='#9C27B0', fg='white', 
                                     activebackground='#7B1FA2',
                                     cursor='hand2', relief=tk.RAISED, bd=4,
                                     padx=30, pady=20, width=20, height=4)
        self.btn_generar.pack(side=tk.LEFT, padx=10)
        
        # Botón de emails
        btn_emails = tk.Button(btn_frame, text="📧 CONFIGURAR\nCORREO", 
                              command=self.preparar_emails,
                              font=fuente_boton,
                              bg='#00BCD4', fg='white',
                              activebackground='#0097A7',
                              cursor='hand2', relief=tk.RAISED, bd=4,
                              padx=30, pady=20, width=20, height=4)
        btn_emails.pack(side=tk.LEFT, padx=10)

        self.btn_borradores = tk.Button(btn_frame, text="📝 CREAR\nBORRADORES\n(Gmail)", 
                      command=self.crear_borradores_gmail,
                      font=fuente_boton,
                      bg='#1976D2', fg='white',
                      activebackground='#1565C0',
                      cursor='hand2', relief=tk.RAISED, bd=4,
                      padx=30, pady=20, width=20, height=4)
        self.btn_borradores.pack(side=tk.LEFT, padx=10)
        
        # Botón para abrir carpeta
        btn_abrir_carpeta = tk.Button(btn_frame, text="📁 ABRIR\nCARPETA\nDE PDFs", 
                              command=self.abrir_carpeta_salida,
                              font=fuente_boton,
                              bg='#4CAF50', fg='white',
                              activebackground='#45a049',
                              cursor='hand2', relief=tk.RAISED, bd=4,
                              padx=30, pady=20, width=20, height=4)
        btn_abrir_carpeta.pack(side=tk.LEFT, padx=10)

        self.btn_enviar_borradores = tk.Button(
            main_frame, text="📤 ENVIAR BORRADORES REVISADOS (Gmail)",
            command=self.enviar_borradores_gmail, font=fuente_boton,
            bg='#D32F2F', fg='white', activebackground='#B71C1C',
            cursor='hand2', padx=20, pady=12,
        )
        self.btn_enviar_borradores.pack(pady=(0, 15))
        
        # Separador
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Área de log
        tk.Label(main_frame, text="► Registro de actividad:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=12, width=80, 
                                                  wrap=tk.WORD, font=fuente_normal)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Botón limpiar log
        btn_limpiar = tk.Button(main_frame, text="🗑️ Limpiar Log", 
                  command=self.limpiar_log, font=fuente_boton,
                  bg='#607D8B', fg='white', activebackground='#455A64',
                  cursor='hand2', relief=tk.RAISED, bd=2, padx=15, pady=8)
        btn_limpiar.pack(pady=10)
        # Agregar mensaje inicial
        self.agregar_log("Sistema iniciado. Seleccione el archivo Excel y la plantilla Word.")
        self.agregar_log("=" * 80)
        
    def seleccionar_excel(self):
        """Abre diálogo para seleccionar archivo Excel"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo_excel.set(archivo)
            self.agregar_log(f"✓ Excel seleccionado: {Path(archivo).name}")
            
            # Abrir ventana de mapeo de campos
            self.abrir_ventana_mapeo()
            
    def seleccionar_plantilla(self):
        """Abre diálogo para seleccionar plantilla Word"""
        # Abrir diálogo en la carpeta de la plantilla actual, o en docs del programa
        actual = Path(self.plantilla_word.get())
        carpeta_docs = ruta_base_app() / "docs"
        if actual.exists():
            carpeta_inicial = str(actual.parent)
        elif carpeta_docs.exists():
            carpeta_inicial = str(carpeta_docs)
        else:
            carpeta_inicial = str(Path.cwd())
        
        archivo = filedialog.askopenfilename(
            title="Seleccionar plantilla Word",
            initialdir=carpeta_inicial,
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.plantilla_word.set(archivo)
            self.agregar_log(f"✓ Plantilla seleccionada: {Path(archivo).name}")
            self._guardar_config()
            
    def seleccionar_carpeta(self):
        """Abre diálogo para seleccionar carpeta de salida"""
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if carpeta:
            self.carpeta_salida.set(carpeta)
            self.agregar_log(f"✓ Carpeta de salida: {carpeta}")
            self._guardar_config()
            
    def agregar_log(self, mensaje):
        """Agrega un mensaje al área de log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def limpiar_log(self):
        """Limpia el área de log"""
        self.log_text.delete(1.0, tk.END)
    
    @staticmethod
    def _parece_fecha(texto: str) -> bool:
        """Detecta valores que son fechas/horas y no texto libre (ej. un país)"""
        texto = texto.strip()
        if not texto:
            return False
        # 2026-06-29 22:46:23.014000, 29/06/2026, 29-06-2026 22:46...
        patrones = (
            r'^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2}(\.\d+)?)?)?$',
            r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}([ T]\d{1,2}:\d{2}(:\d{2})?)?$',
            r'^\d{1,2}:\d{2}(:\d{2})?$',
        )
        return any(re.match(p, texto) for p in patrones)
    
    def abrir_documentacion(self):
        """Abre la guía de uso en el navegador"""
        try:
            webbrowser.open(URL_DOCUMENTACION)
            self.agregar_log("✓ Guía de uso abierta en el navegador")
        except Exception as e:
            self.agregar_log(f"✗ No se pudo abrir la guía: {str(e)}")
            messagebox.showerror(
                "Error",
                "No se pudo abrir el navegador. Abre esta dirección a mano: " + URL_DOCUMENTACION
            )
    
    def _obtener_clave_encriptacion(self) -> bytes:
        """Obtiene una clave de encriptación basada en el usuario del sistema"""
        usuario = os.getenv('USERNAME', 'usuario_desconocido')
        # Crear una clave determinística basada en el usuario
        hash_usuario = hashlib.sha256(usuario.encode()).digest()
        # Fernet requiere una clave en base64
        clave = base64.urlsafe_b64encode(hash_usuario[:32])
        return clave
    
    def _encriptar_texto(self, texto: str) -> str:
        """Encripta un texto"""
        try:
            clave = self._obtener_clave_encriptacion()
            fernet = Fernet(clave)
            texto_encriptado = fernet.encrypt(texto.encode())
            return base64.b64encode(texto_encriptado).decode()
        except Exception:
            # Si falla, devolver el texto en plano
            return texto
    
    def _desencriptar_texto(self, texto_encriptado: str) -> str:
        """Desencripta un texto"""
        try:
            clave = self._obtener_clave_encriptacion()
            fernet = Fernet(clave)
            texto_decodificado = base64.b64decode(texto_encriptado.encode())
            texto_desencriptado = fernet.decrypt(texto_decodificado)
            return texto_desencriptado.decode()
        except Exception:
            # Si falla, devolver el texto como está (probablemente ya estaba en plano)
            return texto_encriptado
    
    def _cargar_config(self) -> dict:
        """Carga la configuración guardada"""
        try:
            if self.ruta_config.exists():
                with open(self.ruta_config, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Desencriptar contraseña si está encriptada
                    if 'gmail_app_password' in config and config['gmail_app_password']:
                        config['gmail_app_password'] = self._desencriptar_texto(config['gmail_app_password'])
                    return config
        except:
            pass
        return {}
    
    def _guardar_config(self):
        """Guarda la configuración actual"""
        try:
            config = {
                "gmail_usuario": self.gmail_usuario.get(),
                "gmail_app_password": self._encriptar_texto(self.gmail_app_password.get()),
                "gmail_nombre_remitente": self.gmail_nombre_remitente.get(),
                "email_asunto": self.email_asunto.get(),
                "email_cuerpo": self.email_cuerpo,
                "plantilla_word": self.plantilla_word.get(),
                "carpeta_salida": self.carpeta_salida.get(),
            }
            with open(self.ruta_config, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.agregar_log(f"✓ Configuración guardada (contraseña encriptada)")
        except Exception as e:
            self.agregar_log(f"⚠ Advertencia: No se pudo guardar configuración: {str(e)}")
        
    def validar_archivos(self):
        """Valida que los archivos necesarios estén seleccionados"""
        if not self.archivo_excel.get():
            messagebox.showerror("Error", "Debe seleccionar un archivo Excel")
            return False
            
        if not self.plantilla_word.get():
            messagebox.showerror("Error", "Debe seleccionar una plantilla Word")
            return False
            
        if not os.path.exists(self.archivo_excel.get()):
            messagebox.showerror("Error", "El archivo Excel no existe")
            return False
            
        if not os.path.exists(self.plantilla_word.get()):
            messagebox.showerror("Error", "La plantilla Word no existe")
            return False
            
        return True
        
    @staticmethod
    def _formatear_duracion(segundos):
        minutos, segundos = divmod(max(0, ceil(segundos)), 60)
        horas, minutos = divmod(minutos, 60)
        if horas:
            return f"{horas} h {minutos:02d} min {segundos:02d} s"
        if minutos:
            return f"{minutos} min {segundos:02d} s"
        return f"{segundos} s"

    def _cerrar_aplicacion(self):
        if self._generando:
            self.ventana_progreso.lift()
            self.ventana_progreso.bell()
            return
        self.root.destroy()

    def _crear_modal_progreso_pdf(self, titulo="Generando consentimientos en PDF"):
        ventana = tk.Toplevel(self.root)
        self.ventana_progreso = ventana
        ventana.title(titulo)
        ventana.transient(self.root)
        ventana.resizable(False, False)
        ventana.protocol("WM_DELETE_WINDOW", lambda: None)
        frame = tk.Frame(ventana, bg='#ffffff', padx=24, pady=22)
        frame.pack(fill=tk.BOTH, expand=True)
        self.progreso_estado = tk.StringVar(master=ventana, value="Preparando generación…")
        self.progreso_detalle = tk.StringVar(master=ventana)
        self.progreso_tiempo = tk.StringVar(master=ventana)
        tk.Label(frame, textvariable=self.progreso_estado, anchor=tk.W,
                 font=('Segoe UI', 12, 'bold'), bg='#ffffff', fg='#333333',
                 wraplength=630).pack(fill=tk.X)
        self.barra_progreso = ttk.Progressbar(frame, mode='determinate', maximum=100, length=630)
        self.barra_progreso.pack(fill=tk.X, pady=(16, 12))
        tk.Label(frame, textvariable=self.progreso_detalle, anchor=tk.W,
                 font=('Segoe UI', 10), bg='#ffffff', fg='#333333').pack(fill=tk.X)
        tk.Label(frame, textvariable=self.progreso_tiempo, anchor=tk.W,
                 font=('Segoe UI', 10), bg='#ffffff', fg='#666666',
                 wraplength=630).pack(fill=tk.X, pady=(8, 0))
        tk.Label(frame, text="Esta ventana se cerrará al terminar.", anchor=tk.W,
                 font=('Segoe UI', 9), bg='#ffffff', fg='#666666').pack(fill=tk.X, pady=(14, 0))
        ventana.update_idletasks()
        x = max(0, self.root.winfo_x() + (self.root.winfo_width() - ventana.winfo_reqwidth()) // 2)
        y = max(0, self.root.winfo_y() + (self.root.winfo_height() - ventana.winfo_reqheight()) // 2)
        ventana.geometry(f"+{x}+{y}")
        ventana.grab_set()

    def _cerrar_modal_progreso_pdf(self):
        self._generando = False
        self.barra_progreso.stop()
        if self.ventana_progreso is not None:
            self.ventana_progreso.grab_release()
            self.ventana_progreso.destroy()
            self.ventana_progreso = None
        self.btn_generar.configure(state=tk.NORMAL)
        if hasattr(self, "btn_borradores"):
            self.btn_borradores.configure(state=tk.NORMAL)
        if hasattr(self, "btn_enviar_borradores"):
            self.btn_enviar_borradores.configure(state=tk.NORMAL)

    def _procesar_eventos_generacion(self):
        """Tk consume los eventos en su propio hilo y mantiene activo el reloj."""
        if not self._generando:
            return
        for _ in range(100):
            try:
                tipo, datos = self._eventos_generacion.get_nowait()
            except Empty:
                break
            if tipo == "log":
                self.agregar_log(datos)
            elif tipo == "total":
                self._iniciar_progreso_pdf(datos)
            elif tipo == "progreso":
                self._actualizar_progreso_pdf(**datos)
            elif tipo == "resultado":
                self._resultado_generacion = datos
            elif tipo == "fin":
                self._cerrar_modal_progreso_pdf()
                if self._resultado_generacion:
                    tipo_aviso, titulo, mensaje = self._resultado_generacion
                    mostrar = messagebox.showerror if tipo_aviso == "error" else messagebox.showinfo
                    mostrar(titulo, mensaje, parent=self.root)
                return
        self._actualizar_progreso_pdf()
        self.root.after(100, self._procesar_eventos_generacion)

    def _iniciar_progreso_pdf(self, total=0):
        self._progreso_inicio = monotonic()
        self._progreso_total = total
        self._progreso_procesados = 0
        self._progreso_generados = 0
        self._progreso_errores = 0
        self._progreso_fin_estimado = None
        self._progreso_terminado = False
        self._actualizar_progreso_pdf(estado="Preparando generación…")

    def _actualizar_progreso_pdf(self, procesados=None, generados=None, errores=None,
                                estado=None, terminado=False):
        ahora = monotonic()
        if procesados is not None and procesados > 0:
            media = max(0, ahora - self._progreso_inicio) / procesados
            self._progreso_fin_estimado = ahora + media * max(0, self._progreso_total - procesados)
        if terminado:
            self._progreso_terminado = True
        if procesados is not None:
            self._progreso_procesados = procesados
        if generados is not None:
            self._progreso_generados = generados
        if errores is not None:
            self._progreso_errores = errores
        total = self._progreso_total
        procesados = self._progreso_procesados
        pendientes = max(0, total - procesados)
        porcentaje = 100 * procesados / total if total else 0
        self.barra_progreso.configure(value=porcentaje)
        if estado:
            self.progreso_estado.set(estado)
        self.progreso_detalle.set(
            f"{procesados} de {total} procesados ({porcentaje:.0f} %) · "
            f"{self._progreso_generados} PDFs generados · "
            f"{pendientes} pendientes · {self._progreso_errores} errores"
            if total else "Leyendo los datos del Excel…"
        )
        transcurrido = max(0, ahora - self._progreso_inicio)
        if self._progreso_terminado:
            self.progreso_tiempo.set(f"Tiempo total: {self._formatear_duracion(transcurrido)}")
            if not total:
                self.progreso_detalle.set("No se ha generado ningún PDF.")
        else:
            if self._progreso_fin_estimado is None:
                restante = "calculando tras el primer PDF…"
            elif not pendientes:
                restante = "finalizando…"
            elif ahora >= self._progreso_fin_estimado:
                restante = "recalculando…"
            else:
                restante = self._formatear_duracion(self._progreso_fin_estimado - ahora)
            self.progreso_tiempo.set(
                f"Transcurrido: {self._formatear_duracion(transcurrido)} · "
                f"Tiempo restante estimado: {restante}"
            )

    def generar_consentimientos(self):
        """Genera los consentimientos desde Excel y plantilla Word"""
        if getattr(self, "_generando", False):
            return
        if not self.validar_archivos():
            return
        
        if not self.mapeo_campos:
            messagebox.showwarning("Advertencia",
                                 "No has configurado el mapeo de campos.\n"
                                 "Selecciona el archivo Excel para configurar el mapeo.")
            return
            
        self._crear_modal_progreso_pdf()
        self._generando = True
        self._iniciar_progreso_pdf()
        self.btn_generar.configure(state=tk.DISABLED)
        self._eventos_generacion = Queue()
        self._resultado_generacion = None
        configuracion = {
            "archivo_excel": self.archivo_excel.get(),
            "plantilla_word": self.plantilla_word.get(),
            "carpeta_salida": self.carpeta_salida.get(),
            "mapeo_campos": dict(self.mapeo_campos),
            "valores_fijos": dict(self.valores_fijos),
        }
        self._hilo_generacion = Thread(
            target=self._generar_consentimientos_trabajo,
            args=(configuracion, self._eventos_generacion),
            daemon=True,
        )
        try:
            self._hilo_generacion.start()
        except Exception as exc:
            self._cerrar_modal_progreso_pdf()
            messagebox.showerror("Error", f"No se pudo iniciar la generación: {exc}", parent=self.root)
            return
        self.root.after(100, self._procesar_eventos_generacion)

    def _generar_consentimientos_trabajo(self, configuracion, eventos):
        """Trabajo bloqueante: comunica resultados por cola, sin acceder a Tk."""
        mapeo_campos = configuracion["mapeo_campos"]
        valores_fijos = configuracion["valores_fijos"]

        def log(mensaje):
            eventos.put(("log", mensaje))

        def progreso(**datos):
            eventos.put(("progreso", datos))

        def avisar(tipo, titulo, mensaje):
            eventos.put(("resultado", (tipo, titulo, mensaje)))

        try:
            from modulos.lector_excel import LectorExcel
            from modulos.generador_word import GeneradorConsentimientos
            from modulos.registro_consentimientos import RegistroConsentimientos
            import subprocess
            import time
            
            log("=" * 80)
            log("Iniciando proceso de generación de consentimientos en PDF...")
            
            # Cerrar cualquier proceso de Word previo
            log("Cerrando procesos previos de Word...")
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                             capture_output=True, timeout=3)
                time.sleep(1)
                log("✓ Procesos de Word cerrados")
            except:
                pass
            
            # Crear carpeta de salida si no existe
            carpeta_salida = Path(configuracion["carpeta_salida"])
            carpeta_salida.mkdir(parents=True, exist_ok=True)
            
            # Invalidar el lote anterior; limpiar sus archivos tras validar el Excel.
            registro_pdfs = RegistroConsentimientos(carpeta_salida)
            registro_pdfs.iniciar([])
            
            # Leer datos del Excel
            log(f"Leyendo datos de: {Path(configuracion['archivo_excel']).name}")
            lector = LectorExcel(configuracion["archivo_excel"])
            datos = lector.leer_datos()
            
            log(f"✓ Se encontraron {len(datos)} registros en el Excel")
            
            if not datos:
                log("✗ Error: No hay datos en el Excel")
                progreso(estado="Sin registros para generar", terminado=True)
                avisar("error", "Error", "No hay datos en el Excel")
                return

            progreso(estado="Eliminando PDFs y borradores anteriores…")
            eliminados = registro_pdfs.iniciar(datos, limpiar_anteriores=True)
            log(f"✓ Eliminados {eliminados} archivos del lote anterior (PDFs y borradores)")
            
            # Aplicar mapeo a los datos
            log(f"Aplicando mapeo de campos: {len(mapeo_campos)} campos")
            if valores_fijos:
                log(f"Aplicando valores fijos: {len(valores_fijos)} campos")
            
            datos_mapeados = []
            paises_descartados = set()
            for registro in datos:
                registro_nuevo = {}
                
                # Primero aplicar mapeo desde Excel
                for campo_plantilla, columna_excel in mapeo_campos.items():
                    if columna_excel in registro:
                        valor = registro[columna_excel]
                        registro_nuevo[campo_plantilla] = valor
                
                # Luego aplicar valores fijos (solo si no se mapearon desde Excel)
                for campo, valor_fijo in valores_fijos.items():
                    if campo not in registro_nuevo:
                        registro_nuevo[campo] = valor_fijo
                
                # El pais nunca puede quedar vacio ni venir de una columna de fecha
                pais = str(registro_nuevo.get('pais_deudor', '')).strip()
                if not pais or self._parece_fecha(pais):
                    if pais:
                        paises_descartados.add(pais)
                    registro_nuevo['pais_deudor'] = PAIS_POR_DEFECTO
                
                # Crear referencia_orden combinando nombre y apellido del alumno
                nombre_alumno = registro_nuevo.get('nombre_alumno', '')
                apellido_alumno = registro_nuevo.get('apellido_alumno', '')
                
                if nombre_alumno and apellido_alumno:
                    referencia = f"{nombre_alumno} {apellido_alumno}_Robotica"
                elif nombre_alumno:
                    referencia = f"{nombre_alumno}_Robotica"
                elif apellido_alumno:
                    referencia = f"{apellido_alumno}_Robotica"
                else:
                    referencia = "Robotica"
                
                registro_nuevo['referencia_orden'] = referencia
                
                # Añadir fecha automáticamente (fecha actual)
                fecha_actual = datetime.now().strftime("%d/%m/%Y")
                registro_nuevo['fecha'] = fecha_actual
                
                # Copiar población a localidad_firma si existe población
                if 'poblacion' in registro_nuevo:
                    registro_nuevo['localidad_firma'] = registro_nuevo['poblacion']
                
                datos_mapeados.append(registro_nuevo)
            
            if paises_descartados:
                ejemplo = next(iter(paises_descartados))
                log(f"⚠ El campo 'País del deudor' recibía valores no válidos (ej: {ejemplo})")
                log(f"  Se ha usado '{PAIS_POR_DEFECTO}'. Revisa el mapeo si no es correcto.")
            
            # Determinar tamaño de lote óptimo
            TAMANO_LOTE = 10
            total_registros = len(datos_mapeados)
            eventos.put(("total", total_registros))
            num_lotes = (total_registros + TAMANO_LOTE - 1) // TAMANO_LOTE
            
            if total_registros > TAMANO_LOTE:
                log(f"📦 Procesamiento en lotes de {TAMANO_LOTE} registros")
                log(f"   Total de lotes: {num_lotes}")
            
            # Generar consentimientos
            generador = GeneradorConsentimientos(configuracion["plantilla_word"], 
                                                str(carpeta_salida))
            
            consentimientos_generados = []
            errores = 0
            
            for lote_num in range(num_lotes):
                inicio_lote = lote_num * TAMANO_LOTE
                fin_lote = min((lote_num + 1) * TAMANO_LOTE, total_registros)
                registros_lote = datos_mapeados[inicio_lote:fin_lote]
                
                if num_lotes > 1:
                    log(f"\n📦 LOTE {lote_num + 1}/{num_lotes} - Registros {inicio_lote + 1} a {fin_lote}")
                
                # Procesar registros del lote
                for i, registro in enumerate(registros_lote, inicio_lote + 1):
                    progreso(estado=f"Generando PDF {i} de {total_registros}…")
                    try:
                        archivo_generado = generador.generar_consentimiento(registro)
                        registro_pdfs.agregar(i, datos[i - 1], archivo_generado)
                        consentimientos_generados.append(archivo_generado)
                        log(f"  [{i}/{total_registros}] ✓ Generado: {Path(archivo_generado).name}")
                    except Exception as e:
                        errores += 1
                        log(f"  [{i}/{total_registros}] ✗ Error: {str(e)}")
                    progreso(
                        procesados=i, generados=len(consentimientos_generados), errores=errores,
                        estado=f"Procesados {i} de {total_registros} registros"
                    )
                
                # Pausa entre lotes para dar tiempo a Word
                if lote_num < num_lotes - 1:  # No pausar después del último lote
                    progreso(estado="Pausa entre lotes: preparando los siguientes PDFs…")
                    log(f"   ⏸ Pausa de 3 segundos antes del siguiente lote...")
                    # Cerrar procesos de Word para limpiar
                    try:
                        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                                     capture_output=True, timeout=2)
                    except:
                        pass
                    time.sleep(3)
            
            registro_pdfs.finalizar()
            if errores:
                log("✗ Borradores bloqueados: vuelve a generar el lote completo sin errores.")

            # Limpieza final de Word
            progreso(estado="Finalizando generación…")
            log("🧹 Limpieza final de procesos de Word...")
            try:
                subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                             capture_output=True, timeout=2)
            except:
                pass
            
            log("=" * 80)
            log(f"RESUMEN: {len(consentimientos_generados)} consentimientos generados correctamente")
            if errores > 0:
                log(f"ADVERTENCIA: {errores} registros con errores")

            progreso(
                estado=(f"Finalizado con {errores} errores. La creación de borradores está bloqueada."
                        if errores else "Todos los PDFs se han generado correctamente"),
                terminado=True,
            )
            
            avisar("info", "Completado", 
                              f"Se generaron {len(consentimientos_generados)} consentimientos en PDF\n"
                              f"Carpeta: {carpeta_salida}")
            
        except ImportError as e:
            progreso(estado="Generación interrumpida: faltan módulos", terminado=True)
            log(f"✗ Error: Módulos no encontrados - {str(e)}")
            avisar("error", "Error", 
                               "Faltan módulos necesarios. Asegúrese de instalar:\n"
                               "pip install openpyxl python-docx")
        except Exception as e:
            progreso(estado="Generación interrumpida por un error", terminado=True)
            log(f"✗ Error: {str(e)}")
            avisar("error", "Error", f"Error al generar consentimientos:\n{str(e)}")
        finally:
            eventos.put(("fin", None))
            
    def preparar_emails(self):
        """Abre ventana para configurar asunto y cuerpo del email"""
        # Crear ventana de configuración
        ventana_config = tk.Toplevel(self.root)
        ventana_config.title("Configurar Emails")
        ventana_config.geometry("800x600")
        ventana_config.resizable(True, True)
        ventana_config.transient(self.root)
        ventana_config.grab_set()
        
        # Centrar ventana
        ventana_config.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (800 // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (600 // 2)
        ventana_config.geometry(f"+{x}+{y}")
        
        # Frame principal
        frame = tk.Frame(ventana_config, padx=20, pady=20, bg='#f5f5f5')
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = tk.Label(frame, 
                        text="📧 Configurar Asunto y Cuerpo del Email", 
                        font=('Arial', 16, 'bold'),
                        bg='#f5f5f5', fg='#333')
        titulo.pack(pady=(0, 20))
        
        instruccion = tk.Label(frame,
                             text="Configura el texto y revisa la lista de adjuntos antes de crear los borradores en Gmail.",
                             font=('Arial', 10),
                             bg='#f5f5f5', fg='#555')
        instruccion.pack(pady=(0, 15))
        
        # Sección Asunto
        tk.Label(frame, text="Asunto del email:", font=('Arial', 11, 'bold'), 
                bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(10, 5))
        
        entry_asunto = ttk.Entry(frame, font=('Arial', 11), width=80)
        entry_asunto.insert(0, self.email_asunto.get())
        entry_asunto.pack(fill=tk.X, pady=(0, 20))
        
        # Sección Cuerpo
        tk.Label(frame, text="Cuerpo del email:", font=('Arial', 11, 'bold'),
                bg='#f5f5f5', fg='#333').pack(anchor=tk.W, pady=(10, 5))
        
        tk.Label(frame, text="Usa {nombre} para personalizar con el nombre de cada destinatario",
                font=('Arial', 9), bg='#f5f5f5', fg='#666').pack(anchor=tk.W, pady=(0, 5))
        
        # Texto por defecto si no hay nada guardado
        texto_defecto = """Estimado/a {nombre},

Adjunto encontrará el Mandato de domiciliación SEPA necesario para gestionar los cobros de su servicio de forma automática a través de su cuenta bancaria.

Pasos para completar el proceso:

1. Descargue y revise el documento adjunto

2. Firme el documento (se acepta firma manuscrita escaneada o firma digital).

3. Responda a este mismo correo adjuntando el documento firmado.

Este trámite es indispensable para asegurar que sus pagos se procesen correctamente y sin interrupciones.

Si tiene alguna pregunta o inquietud, no dude en contactarnos respondiendo a este mensaje.

Saludos cordiales,

Isabel Garcia Rincon
Gfg Kids"""
        
        text_cuerpo = scrolledtext.ScrolledText(frame, height=14, width=80, 
                                               wrap=tk.WORD, font=('Arial', 10))
        text_cuerpo.insert(1.0, self.email_cuerpo if self.email_cuerpo else texto_defecto)
        text_cuerpo.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Botones
        btn_frame = tk.Frame(frame, bg='#f5f5f5')
        btn_frame.pack(fill=tk.X, pady=10)
        
        def generar_borradores_emails():
            asunto = entry_asunto.get().strip()
            cuerpo = text_cuerpo.get(1.0, tk.END).strip()
            
            if not asunto or not cuerpo:
                messagebox.showwarning("Datos incompletos", "Introduce asunto y cuerpo")
                return
            
            # Guardar en variables
            self.email_asunto.set(asunto)
            self.email_cuerpo = cuerpo
            
            # Guardar en archivo de configuración
            self._guardar_config()
            
            ventana_config.destroy()
            
            # Generar borradores automáticamente
            carpeta_salida = Path(self.carpeta_salida.get())
            
            if not carpeta_salida.exists():
                messagebox.showerror("Error", "No hay consentimientos generados")
                return
            
            try:
                from modulos.preparador_emails import PreparadorEmails
                
                self.agregar_log("=" * 80)
                self.agregar_log("Preparando la lista de correos y adjuntos...")
                
                preparador = PreparadorEmails(str(carpeta_salida))
                resultado = preparador.crear_borradores_email(
                    self.archivo_excel.get(),
                    asunto_base=asunto,
                    cuerpo_personalizado=cuerpo
                )
                
                self.agregar_log(f"✓ Se prepararon {len(resultado)} correos para revisión")
                self.agregar_log(f"✓ Archivo 'emails_para_enviar.txt' generado")
                self.agregar_log(f"✓ Archivo 'emails_lista.csv' generado")
                
                messagebox.showinfo("Completado", 
                                  f"Se prepararon {len(resultado)} correos para revisión\n"
                                  "Revisa la lista y pulsa Crear borradores (Gmail)")
                
            except ImportError:
                self.agregar_log("✗ Error: Módulo preparador_emails no encontrado")
                messagebox.showerror("Error", "Módulo preparador_emails no encontrado")
            except Exception as e:
                self.agregar_log(f"✗ Error: {str(e)}")
                messagebox.showerror("Error", f"Error al preparar emails:\n{str(e)}")
        
        btn_generar = tk.Button(btn_frame, text="✓ Guardar y preparar revisión",
                              command=generar_borradores_emails,
                              font=('Arial', 12, 'bold'),
                              bg='#4CAF50', fg='white',
                              activebackground='#45a049',
                              cursor='hand2', relief=tk.RAISED, bd=3,
                              padx=30, pady=10)
        btn_generar.pack(side=tk.LEFT, padx=10)
        
        btn_cancelar = tk.Button(btn_frame, text="✗ Cancelar",
                               command=ventana_config.destroy,
                               font=('Arial', 12, 'bold'),
                               bg='#f44336', fg='white',
                               activebackground='#da190b',
                               cursor='hand2', relief=tk.RAISED, bd=3,
                               padx=30, pady=10)
        btn_cancelar.pack(side=tk.LEFT, padx=10)

    def crear_borradores_gmail(self):
        """Crea borradores revisables en Gmail, sin posibilidad de envío directo."""
        if self._generando:
            return
        usuario = self.gmail_usuario.get().strip()
        clave = self.gmail_app_password.get().strip()
        if not usuario or not clave:
            messagebox.showwarning("Datos incompletos", "Indica la cuenta Gmail y su contraseña de aplicación.")
            return
        if not self.archivo_excel.get() or not Path(self.carpeta_salida.get()).exists():
            messagebox.showerror("Faltan datos", "Selecciona el Excel y genera primero los PDFs.")
            return
        self._guardar_config()
        configuracion = {
            "ruta_excel": self.archivo_excel.get(), "gmail_usuario": usuario,
            "gmail_app_password": clave,
            "nombre_remitente": self.gmail_nombre_remitente.get().strip() or None,
            "asunto_base": self.email_asunto.get(), "cuerpo_personalizado": self.email_cuerpo,
        }
        carpeta = self.carpeta_salida.get()
        self._crear_modal_progreso_pdf(titulo="Creando borradores en Gmail")
        self._generando = True
        self.btn_generar.configure(state=tk.DISABLED)
        self.btn_borradores.configure(state=tk.DISABLED)
        self.progreso_estado.set("Verificando destinatarios y PDFs…")
        self.progreso_detalle.set("Los mensajes se guardarán en Borradores de Gmail para que los revises.")
        self.progreso_tiempo.set("Transcurrido: 0 s")
        self.barra_progreso.configure(mode="indeterminate")
        self.barra_progreso.start(15)
        self._inicio_borradores = monotonic()
        self._eventos_borradores = Queue()
        self._resultado_borradores = None
        self._error_borradores = None
        self._hilo_borradores = Thread(
            target=self._crear_borradores_trabajo,
            args=(carpeta, configuracion, self._eventos_borradores), daemon=True,
        )
        try:
            self._hilo_borradores.start()
        except Exception as exc:
            self._cerrar_modal_progreso_pdf()
            messagebox.showerror("Error", f"No se pudo iniciar la creación de borradores: {exc}")
            return
        self.root.after(100, self._procesar_eventos_borradores)

    @staticmethod
    def _crear_borradores_trabajo(carpeta, configuracion, eventos):
        try:
            from modulos.preparador_emails import PreparadorEmails
            resultado = PreparadorEmails(carpeta).crear_borradores_gmail(
                **configuracion,
                progreso=lambda *datos: eventos.put(("progreso", datos)),
            )
            eventos.put(("resultado", resultado))
        except Exception as exc:
            eventos.put(("error", str(exc)))
        finally:
            eventos.put(("fin", None))

    def _procesar_eventos_borradores(self):
        if not self._generando:
            return
        for _ in range(100):
            try:
                tipo, datos = self._eventos_borradores.get_nowait()
            except Empty:
                break
            if tipo == "progreso":
                procesados, total, creados, existentes = datos
                self.barra_progreso.stop()
                self.barra_progreso.configure(mode="determinate", value=100 * procesados / total if total else 0)
                self.progreso_estado.set("Conectando con Gmail…" if not procesados else "Guardando borradores en Gmail…")
                self.progreso_detalle.set(
                    f"{procesados} de {total} revisados · {creados} creados · "
                    f"{existentes} ya existentes · {total - procesados} pendientes"
                )
            elif tipo == "resultado":
                self._resultado_borradores = datos
            elif tipo == "error":
                self._error_borradores = datos
            elif tipo == "fin":
                self._cerrar_modal_progreso_pdf()
                if self._error_borradores:
                    self.agregar_log(f"✗ Error al crear borradores: {self._error_borradores}")
                    messagebox.showerror("Error al crear borradores", self._error_borradores, parent=self.root)
                else:
                    resultado = self._resultado_borradores
                    resumen = (f"Creados: {resultado['creados']}\n"
                               f"Ya existentes en Borradores: {resultado['existentes']}\n"
                               f"Sin confirmar: {len(resultado['fallidos'])}\n"
                               f"Pendientes: {resultado['pendientes']}")
                    self.agregar_log(resumen.replace("\n", " · "))
                    for fallo in resultado["fallidos"]:
                        self.agregar_log(f"  #{fallo['numero']} {fallo['email']}: {fallo['motivo']}")
                    aviso = messagebox.showwarning if resultado["fallidos"] else messagebox.showinfo
                    aviso("Borradores en Gmail", resumen + "\n\nAbre Gmail → Borradores. "
                          "Comprueba cada destinatario y PDF; después vuelve a la app y pulsa "
                          "Enviar borradores revisados.\nEl programa no ha enviado ningún correo.", parent=self.root)
                return
        self.progreso_tiempo.set(f"Transcurrido: {self._formatear_duracion(monotonic() - self._inicio_borradores)}")
        self.root.after(100, self._procesar_eventos_borradores)

    def enviar_borradores_gmail(self):
        if self._generando:
            return
        usuario = self.gmail_usuario.get().strip()
        clave = self.gmail_app_password.get().strip()
        if not usuario or not clave or not self.archivo_excel.get():
            messagebox.showwarning("Faltan datos", "Selecciona el Excel e indica la cuenta Gmail y su contraseña de aplicación.")
            return
        configuracion = {
            "ruta_excel": self.archivo_excel.get(), "gmail_usuario": usuario,
            "gmail_app_password": clave, "nombre_remitente": self.gmail_nombre_remitente.get().strip() or None,
            "asunto_base": self.email_asunto.get(), "cuerpo_personalizado": self.email_cuerpo,
        }
        self._iniciar_operacion_envio(self.carpeta_salida.get(), configuracion)

    def _iniciar_operacion_envio(self, carpeta, configuracion, plan=None):
        self._crear_modal_progreso_pdf(
            titulo="Comprobando borradores en Gmail" if plan is None else "Enviando borradores revisados"
        )
        self._generando = True
        for boton in (self.btn_generar, self.btn_borradores, self.btn_enviar_borradores):
            boton.configure(state=tk.DISABLED)
        self.progreso_estado.set("Leyendo y verificando los borradores…" if plan is None else "Verificando de nuevo antes de enviar…")
        self.progreso_detalle.set("Solo se procesan los borradores correspondientes al lote actual.")
        self.progreso_tiempo.set("Transcurrido: 0 s")
        self.barra_progreso.configure(mode="indeterminate")
        self.barra_progreso.start(15)
        self._inicio_envio = monotonic()
        self._eventos_envio = Queue()
        self._resultado_envio = None
        self._error_envio = None
        self._contexto_envio = (carpeta, configuracion, plan is None)
        self._hilo_envio = Thread(target=self._trabajo_envio,
                                  args=(carpeta, configuracion, plan, self._eventos_envio), daemon=True)
        try:
            self._hilo_envio.start()
        except Exception as exc:
            self._cerrar_modal_progreso_pdf()
            messagebox.showerror("Error", str(exc), parent=self.root)
            return
        self.root.after(100, self._procesar_eventos_envio)

    @staticmethod
    def _trabajo_envio(carpeta, configuracion, plan, eventos):
        try:
            from modulos.envio_borradores import EnviadorBorradores
            servicio = EnviadorBorradores(carpeta)
            progreso = lambda *datos: eventos.put(("progreso", datos))
            resultado = (servicio.preparar_envio(configuracion, progreso) if plan is None
                         else servicio.enviar(plan, configuracion, progreso))
            eventos.put(("resultado", resultado))
        except Exception as exc:
            eventos.put(("error", str(exc)))
        finally:
            eventos.put(("fin", None))

    def _procesar_eventos_envio(self):
        if not self._generando:
            return
        carpeta, configuracion, es_revision = self._contexto_envio
        for _ in range(100):
            try:
                tipo, datos = self._eventos_envio.get_nowait()
            except Empty:
                break
            if tipo == "progreso":
                actual, total = datos
                self.barra_progreso.stop()
                self.barra_progreso.configure(mode="determinate", value=100 * actual / total if total else 0)
                self.progreso_estado.set("Comprobando borradores…" if es_revision else "Enviando borradores revisados…")
                self.progreso_detalle.set(f"{actual} de {total} procesados")
            elif tipo == "resultado":
                self._resultado_envio = datos
            elif tipo == "error":
                self._error_envio = datos
            elif tipo == "fin":
                self._cerrar_modal_progreso_pdf()
                if self._error_envio:
                    self.agregar_log(f"✗ Borradores: {self._error_envio}")
                    messagebox.showerror("No se pudo completar la operación", self._error_envio, parent=self.root)
                elif es_revision:
                    self._mostrar_confirmacion_borradores(self._resultado_envio, carpeta, configuracion)
                else:
                    r = self._resultado_envio
                    texto = (f"Enviados: {r['enviados']}\nYa enviados anteriormente: {r['omitidos']}\n"
                             f"Pendientes: {r['pendientes']}")
                    detalles = r["errores"] + r["avisos"]
                    if detalles:
                        texto += "\n\n" + "\n".join(detalles)
                    self.agregar_log(texto)
                    mostrar = messagebox.showwarning if detalles else messagebox.showinfo
                    mostrar("Resultado del envío", texto, parent=self.root)
                return
        self.progreso_tiempo.set(f"Transcurrido: {self._formatear_duracion(monotonic() - self._inicio_envio)}")
        self.root.after(100, self._procesar_eventos_envio)

    def _mostrar_confirmacion_borradores(self, plan, carpeta, configuracion):
        if not plan["mensajes"]:
            messagebox.showinfo("Sin envíos pendientes", "Los correos de este lote ya constan como enviados.", parent=self.root)
            return
        ventana = tk.Toplevel(self.root)
        self.ventana_progreso = ventana
        self._generando = True
        ventana.title("Confirmar envío de borradores revisados")
        ventana.geometry("900x480")
        ventana.transient(self.root)
        ventana.grab_set()
        marco = ttk.Frame(ventana, padding=18)
        marco.pack(fill=tk.BOTH, expand=True)
        ttk.Label(marco, text=f"Se enviarán {len(plan['mensajes'])} borradores desde {plan['cuenta']}.",
                  font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W)
        ttk.Label(marco, text="Confirma después de revisar sus destinatarios y PDFs en Gmail.").pack(anchor=tk.W, pady=(6, 12))
        tabla = ttk.Treeview(marco, columns=("destino", "asunto", "pdf"), show="headings", height=12)
        for clave, titulo, ancho in (("destino", "Destinatario", 220), ("asunto", "Asunto en Gmail", 240), ("pdf", "PDF adjunto", 360)):
            tabla.heading(clave, text=titulo)
            tabla.column(clave, width=ancho)
        barra = ttk.Scrollbar(marco, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=barra.set)
        barra.pack(side=tk.RIGHT, fill=tk.Y)
        tabla.pack(fill=tk.BOTH, expand=True)
        for item in plan["mensajes"]:
            tabla.insert("", tk.END, values=(item["email"], item["asunto"], item["archivo"]))

        def cerrar():
            ventana.grab_release()
            ventana.destroy()
            self.ventana_progreso = None
            self._generando = False

        def confirmar():
            cerrar()
            self._iniciar_operacion_envio(carpeta, configuracion, plan)

        ventana.protocol("WM_DELETE_WINDOW", cerrar)
        acciones = ttk.Frame(marco)
        acciones.pack(fill=tk.X, pady=(15, 0))
        ttk.Button(acciones, text="Cancelar", command=cerrar).pack(side=tk.LEFT)
        ttk.Button(acciones, text=f"Enviar los {len(plan['mensajes'])} borradores revisados",
                   command=confirmar).pack(side=tk.RIGHT)

    def abrir_carpeta_salida(self):
        """Abre la carpeta donde se guardaron los PDFs"""
        carpeta_salida = Path(self.carpeta_salida.get())
        
        if not carpeta_salida.exists():
            messagebox.showwarning("Advertencia", 
                                 "La carpeta aún no existe.\n"
                                 "Genere consentimientos primero.")
            return
        
        try:
            # Abrir carpeta en el explorador de Windows
            os.startfile(str(carpeta_salida))
            self.agregar_log(f"✓ Carpeta abierta: {carpeta_salida}")
        except Exception as e:
            self.agregar_log(f"✗ Error al abrir carpeta: {str(e)}")
            messagebox.showerror("Error", f"No se pudo abrir la carpeta:\n{str(e)}")
    
    def buscar_datos_codigo_postal(self, codigo_postal, ventana_padre=None):
        """Busca población, provincia y país basándose en el código postal"""
        try:
            # Detectar país por prefijo del código postal
            pais_code = 'ES'  # Por defecto España
            cp_limpio = codigo_postal.strip()
            
            # API gratuita de Zippopotam
            url = f"http://api.zippopotam.us/{pais_code}/{cp_limpio}"
            
            with urllib.request.urlopen(url, timeout=3) as response:
                data = json.loads(response.read().decode())
                
                if 'places' in data and len(data['places']) > 0:
                    lugar = data['places'][0]
                    poblacion = lugar.get('place name', '')
                    provincia = lugar.get('state', '')
                    pais = data.get('country', 'España')
                    
                    # Rellenar los campos automáticamente
                    if 'poblacion' in self.textboxes_fijos and not self.textboxes_fijos['poblacion'].get():
                        self.textboxes_fijos['poblacion'].delete(0, tk.END)
                        self.textboxes_fijos['poblacion'].insert(0, poblacion)
                    
                    if 'provincia' in self.textboxes_fijos and not self.textboxes_fijos['provincia'].get():
                        self.textboxes_fijos['provincia'].delete(0, tk.END)
                        self.textboxes_fijos['provincia'].insert(0, provincia)
                    
                    if 'pais_deudor' in self.textboxes_fijos and not self.textboxes_fijos['pais_deudor'].get():
                        self.textboxes_fijos['pais_deudor'].delete(0, tk.END)
                        self.textboxes_fijos['pais_deudor'].insert(0, pais)
                    
                    # Mostrar mensaje de éxito
                    if ventana_padre:
                        ventana_padre.title(f"Mapeo de Campos - ✓ CP {cp_limpio}: {poblacion}, {provincia}")
        
        except urllib.error.URLError:
            # Sin conexión o código postal no encontrado, no hacer nada
            pass
        except Exception as e:
            # Error silencioso, no molestar al usuario
            pass
    
    def buscar_datos_poblacion(self, poblacion, ventana_padre=None):
        """Busca código postal, provincia y país basándose en el nombre de la población"""
        try:
            # Usar Nominatim API de OpenStreetMap (gratuita)
            poblacion_encoded = urllib.parse.quote(poblacion)
            url = f"https://nominatim.openstreetmap.org/search?city={poblacion_encoded}&country=Spain&format=json&limit=1"
            
            req = urllib.request.Request(url)
            req.add_header('User-Agent', 'SEPAS-App/1.0')
            
            with urllib.request.urlopen(req, timeout=3) as response:
                data = json.loads(response.read().decode())
                
                if data and len(data) > 0:
                    lugar = data[0]
                    lat = lugar.get('lat')
                    lon = lugar.get('lon')
                    
                    # Ahora buscar detalles con reverse geocoding
                    url_reverse = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json"
                    req_reverse = urllib.request.Request(url_reverse)
                    req_reverse.add_header('User-Agent', 'SEPAS-App/1.0')
                    
                    with urllib.request.urlopen(req_reverse, timeout=3) as response_reverse:
                        data_reverse = json.loads(response_reverse.read().decode())
                        address = data_reverse.get('address', {})
                        
                        codigo_postal = address.get('postcode', '')
                        provincia = address.get('state', address.get('province', ''))
                        pais = address.get('country', 'España')
                        
                        # Rellenar los campos automáticamente
                        if 'codigo_postal' in self.textboxes_fijos and not self.textboxes_fijos['codigo_postal'].get():
                            if codigo_postal:
                                self.textboxes_fijos['codigo_postal'].delete(0, tk.END)
                                self.textboxes_fijos['codigo_postal'].insert(0, codigo_postal)
                        
                        if 'provincia' in self.textboxes_fijos and not self.textboxes_fijos['provincia'].get():
                            self.textboxes_fijos['provincia'].delete(0, tk.END)
                            self.textboxes_fijos['provincia'].insert(0, provincia)
                        
                        if 'pais_deudor' in self.textboxes_fijos and not self.textboxes_fijos['pais_deudor'].get():
                            self.textboxes_fijos['pais_deudor'].delete(0, tk.END)
                            self.textboxes_fijos['pais_deudor'].insert(0, pais)
                        
                        # Mostrar mensaje de éxito
                        if ventana_padre:
                            ventana_padre.title(f"Mapeo de Campos - ✓ {poblacion}: CP {codigo_postal}, {provincia}")
        
        except urllib.error.URLError:
            # Sin conexión o población no encontrada, no hacer nada
            pass
        except Exception as e:
            # Error silencioso, no molestar al usuario
            pass
    
    def abrir_ventana_mapeo(self):
        """Abre ventana para mapear columnas del Excel a campos de la plantilla"""
        if not self.archivo_excel.get():
            return
        
        try:
            from modulos.lector_excel import LectorExcel
            
            # Leer columnas del Excel
            lector = LectorExcel(self.archivo_excel.get())
            columnas_excel = lector.obtener_columnas()
            
            if not columnas_excel:
                self.agregar_log("✗ No se encontraron columnas en el Excel")
                messagebox.showerror("Error", "No se encontraron columnas en el Excel")
                return
            
            # Campos necesarios para SEPA
            # Nota: fecha y localidad_firma se rellenan automáticamente
            campos_necesarios = [
                ("nombre_alumno", "NOMBRE del alumno"),
                ("apellido_alumno", "APELLIDO del alumno"),
                ("nombre_deudor", "Nombre del deudor/es (titular de la cuenta)"),
                ("direccion_deudor", "Dirección del deudor"),
                ("codigo_postal", "Código postal"),
                ("poblacion", "Población (se usará también como localidad de firma)"),
                ("provincia", "Provincia"),
                ("pais_deudor", "País del deudor"),
                ("iban", "🏦 IBAN (número de cuenta bancaria)"),
                ("swift_bic", "🏦 Swift BIC del banco"),
                ("identificador_acreedor", "Identificador del acreedor"),
                ("nombre_acreedor", "Nombre del acreedor"),
                ("direccion_acreedor", "Dirección del acreedor"),
                ("firma", "Firma del deudor"),
            ]
            
            # Crear ventana de mapeo
            ventana_mapeo = tk.Toplevel(self.root)
            ventana_mapeo.title("Mapeo de Campos")
            ventana_mapeo.geometry("1200x700")
            ventana_mapeo.resizable(True, True)
            ventana_mapeo.transient(self.root)
            ventana_mapeo.grab_set()
            
            # Centrar la ventana respecto a la ventana principal
            ventana_mapeo.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (1200 // 2)
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (700 // 2)
            ventana_mapeo.geometry(f"+{x}+{y}")
            
            # Frame principal
            frame = tk.Frame(ventana_mapeo, padx=20, pady=20, bg='#f5f5f5')
            frame.pack(fill=tk.BOTH, expand=True)
            
            # Título
            titulo = tk.Label(frame, 
                            text="🔗 Configuración de Campos", 
                            font=('Arial', 16, 'bold'),
                            bg='#f5f5f5', fg='#333')
            titulo.pack(pady=(0, 10))
            
            instruccion = tk.Label(frame,
                                 text=f"Columnas encontradas en Excel: {len(columnas_excel)}\n"
                                      "Para cada campo, selecciona la columna del Excel correspondiente:",
                                 font=('Arial', 10),
                                 bg='#f5f5f5', fg='#555',
                                 wraplength=750)
            instruccion.pack(pady=(0, 20))
            
            # Frame con scroll para los mapeos
            canvas_frame = tk.Frame(frame, bg='#f5f5f5')
            canvas_frame.pack(fill=tk.BOTH, expand=True)
            
            canvas = tk.Canvas(canvas_frame, bg='#f5f5f5', highlightthickness=0)
            scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg='#f5f5f5')
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=1150)
            canvas.configure(yscrollcommand=scrollbar.set)
            
            # Función para scroll con rueda del ratón
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            # Bind scroll del ratón al canvas y al frame scrollable
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
            
            # Limpiar el bind cuando se cierre la ventana
            def on_close():
                canvas.unbind_all("<MouseWheel>")
                ventana_mapeo.destroy()
            
            ventana_mapeo.protocol("WM_DELETE_WINDOW", on_close)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Crear comboboxes para cada campo necesario
            self.comboboxes_mapeo = {}
            self.textboxes_fijos = {}  # Para valores fijos opcionales
            columnas_con_vacio = ["(No mapear - dejar vacío)"] + columnas_excel
            
            # Campos que pueden tener valores fijos
            campos_con_valor_fijo = ['codigo_postal', 'poblacion', 'provincia', 'pais_deudor']
            
            for i, (campo_interno, descripcion) in enumerate(campos_necesarios):
                # Frame para cada fila - ocupar todo el ancho
                fila_frame = tk.Frame(scrollable_frame, bg='white', relief=tk.RAISED, bd=1)
                fila_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
                
                # Label con la pregunta
                label_pregunta = tk.Label(fila_frame,
                                        text=f"{descripcion}:",
                                        font=('Arial', 10),
                                        bg='white', fg='#333',
                                        width=45, anchor='w', padx=10)
                label_pregunta.pack(side=tk.LEFT, pady=10)
                
                # Flecha
                tk.Label(fila_frame, text="→", font=('Arial', 14), bg='white').pack(side=tk.LEFT, padx=10)
                
                # Combobox para seleccionar columna del Excel
                ancho_combo = 30 if campo_interno in campos_con_valor_fijo else 45
                combo = ttk.Combobox(fila_frame,
                                   values=columnas_con_vacio,
                                   state='readonly',
                                   font=('Arial', 10),
                                   width=ancho_combo)
                combo.set("(No mapear - dejar vacío)")
                combo.pack(side=tk.LEFT, pady=10, padx=5)
                
                # Si es un campo con valor fijo opcional, añadir Entry
                if campo_interno in campos_con_valor_fijo:
                    tk.Label(fila_frame, text="o valor fijo:", font=('Arial', 9), 
                           bg='white', fg='#666').pack(side=tk.LEFT, padx=5)
                    entry_fijo = ttk.Entry(fila_frame, font=('Arial', 10), width=20)
                    entry_fijo.pack(side=tk.LEFT, pady=10, padx=5)
                    self.textboxes_fijos[campo_interno] = entry_fijo
                    
                    # Si es código postal, añadir evento para buscar automáticamente
                    if campo_interno == 'codigo_postal':
                        def buscar_por_cp(event):
                            cp = self.textboxes_fijos['codigo_postal'].get().strip()
                            if cp and len(cp) >= 4:
                                self.buscar_datos_codigo_postal(cp, ventana_mapeo)
                        entry_fijo.bind('<FocusOut>', buscar_por_cp)
                        entry_fijo.bind('<Return>', buscar_por_cp)
                    
                    # Si es población, añadir evento para buscar automáticamente
                    elif campo_interno == 'poblacion':
                        def buscar_por_poblacion(event):
                            pob = self.textboxes_fijos['poblacion'].get().strip()
                            if pob and len(pob) >= 3:
                                self.buscar_datos_poblacion(pob, ventana_mapeo)
                        entry_fijo.bind('<FocusOut>', buscar_por_poblacion)
                        entry_fijo.bind('<Return>', buscar_por_poblacion)
                
                # Intentar auto-mapeo inteligente por similitud de nombre
                campo_lower = campo_interno.lower()
                
                # Palabras clave para cada campo
                keywords_map = {
                    'nombre_alumno': ['nombre', 'nom', 'name', 'alumne', 'alumno', 'estudiante', 'student'],
                    'apellido_alumno': ['apellido', 'cognom', 'surname', 'lastname', 'apellidos', 'cognoms'],
                    'nombre_deudor': ['padre', 'madre', 'tutor', 'titular', 'deudor', 'debtor', 'responsable'],
                    'direccion_deudor': ['direccion', 'dirección', 'adreça', 'address', 'calle', 'carrer'],
                    'codigo_postal': ['codigo postal', 'cp', 'postal', 'codi postal', 'zip'],
                    'poblacion': ['poblacion', 'població', 'poblacio', 'ciudad', 'city', 'localidad', 'localitat'],
                    'provincia': ['provincia', 'prov', 'province'],
                    'pais_deudor': ['pais', 'país', 'country', 'nacionalidad'],
                    'swift_bic': ['swift', 'bic', 'banco', 'bank', 'entidad'],
                    'iban': ['iban', 'cuenta', 'account', 'numero cuenta', 'número cuenta', 'compte', 'bancaria', 'bancari', 'numero de compte', 'número de cuenta'],
                    'identificador_acreedor': ['identificador', 'id acreedor', 'creditor'],
                    'nombre_acreedor': ['acreedor', 'empresa', 'organizacion', 'creditor'],
                    'direccion_acreedor': ['direccion acreedor', 'domicilio social'],
                    'firma': ['firma', 'signature', 'signatura']
                }
                
                # Buscar coincidencias usando las palabras clave
                mejor_coincidencia = None
                mejor_puntuacion = 0
                
                for columna in columnas_excel:
                    columna_lower = str(columna).lower()
                    
                    # Las columnas de marca temporal no son datos del deudor
                    if any(marca in columna_lower for marca in COLUMNAS_MARCA_TEMPORAL):
                        continue
                    
                    puntuacion = 0
                    
                    # Verificar palabras clave específicas del campo
                    if campo_interno in keywords_map:
                        for keyword in keywords_map[campo_interno]:
                            if keyword in columna_lower:
                                puntuacion += len(keyword)
                    
                    # Verificar si el nombre del campo está en la columna
                    if campo_lower.replace('_', ' ') in columna_lower:
                        puntuacion += 10
                    
                    if puntuacion > mejor_puntuacion:
                        mejor_puntuacion = puntuacion
                        mejor_coincidencia = columna
                
                if mejor_coincidencia and mejor_puntuacion > 0:
                    combo.set(mejor_coincidencia)
                
                self.comboboxes_mapeo[campo_interno] = combo
            
            # Botones
            btn_frame = tk.Frame(frame, bg='#f5f5f5')
            btn_frame.pack(pady=20)
            
            def guardar_mapeo():
                # Limpiar bind del scroll antes de cerrar
                canvas.unbind_all("<MouseWheel>")
                
                # Guardar el mapeo y valores fijos
                self.mapeo_campos = {}
                self.valores_fijos = {}
                
                for campo_interno, combo in self.comboboxes_mapeo.items():
                    valor = combo.get()
                    if valor != "(No mapear - dejar vacío)":
                        self.mapeo_campos[campo_interno] = valor
                    # Si no está mapeado pero tiene valor fijo, guardarlo
                    elif campo_interno in self.textboxes_fijos:
                        valor_fijo = self.textboxes_fijos[campo_interno].get().strip()
                        if valor_fijo:
                            self.valores_fijos[campo_interno] = valor_fijo
                
                ventana_mapeo.destroy()
                self.agregar_log(f"✓ Mapeo configurado: {len(self.mapeo_campos)} campos mapeados")
                if self.valores_fijos:
                    self.agregar_log(f"✓ Valores fijos: {len(self.valores_fijos)} campos con valor fijo")
                    for campo, valor in self.valores_fijos.items():
                        self.agregar_log(f"    - {campo}: {valor}")
                campos_texto = ", ".join([f"{k}" for k in self.mapeo_campos.keys()])
                if campos_texto:
                    self.agregar_log(f"  Campos mapeados: {campos_texto}")
            
            btn_guardar = tk.Button(btn_frame, text="✓ Guardar Configuración",
                                  command=guardar_mapeo,
                                  font=('Arial', 12, 'bold'),
                                  bg='#4CAF50', fg='white',
                                  activebackground='#45a049',
                                  cursor='hand2', relief=tk.RAISED, bd=3,
                                  padx=30, pady=10)
            btn_guardar.pack(side=tk.LEFT, padx=10)
            
            btn_cancelar = tk.Button(btn_frame, text="✗ Cancelar",
                                   command=ventana_mapeo.destroy,
                                   font=('Arial', 12, 'bold'),
                                   bg='#f44336', fg='white',
                                   activebackground='#da190b',
                                   cursor='hand2', relief=tk.RAISED, bd=3,
                                   padx=30, pady=10)
            btn_cancelar.pack(side=tk.LEFT, padx=10)
            
        except Exception as e:
            self.agregar_log(f"✗ Error al crear ventana de mapeo: {str(e)}")
            messagebox.showerror("Error", f"Error al crear ventana de mapeo:\n{str(e)}")


def main():
    """Función principal"""
    root = tk.Tk()
    app = AplicacionConsentimientos(root)
    root.mainloop()


if __name__ == "__main__":
    main()
