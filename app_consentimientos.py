"""
Sistema de Generación de Consentimientos
Aplicación para leer datos de Excel, generar consentimientos desde plantilla Word
y prepararlos para envío por correo electrónico.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from pathlib import Path
from datetime import datetime
import urllib.request
import json


class AplicacionConsentimientos:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Consentimientos - SEPAS")
        self.root.geometry("950x750")
        self.root.state('zoomed')  # Abrir maximizado en Windows
        self.root.resizable(True, True)
        
        # Color de fondo
        self.root.configure(bg='#f5f5f5')
        
        # Variables
        self.archivo_excel = tk.StringVar()
        # Buscar plantilla en carpeta docs del software
        ruta_software = Path(__file__).parent
        ruta_plantilla_default = ruta_software / "docs" / "plantilla_domiciliacion_sepa.docx"
        self.plantilla_word = tk.StringVar(value=str(ruta_plantilla_default))
        self.carpeta_salida = tk.StringVar(value=str(Path.cwd() / "consentimientos_generados"))
        
        # Mapeo de columnas Excel a campos de la plantilla
        self.mapeo_campos = {}
        self.valores_fijos = {}  # Valores fijos para campos no mapeados
        
        # Crear interfaz
        self.crear_interfaz()
        
    def crear_interfaz(self):
        """Crea la interfaz gráfica de la aplicación"""
        
        # Frame principal con padding y color de fondo
        main_frame = tk.Frame(self.root, bg='#f5f5f5', padx=20, pady=10)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar peso de filas y columnas
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Configurar fuentes más grandes
        fuente_titulo = ('Arial', 24, 'bold')
        fuente_seccion = ('Arial', 13, 'bold')
        fuente_normal = ('Arial', 11)
        fuente_boton = ('Arial', 14, 'bold')
        fuente_boton_grande = ('Arial', 16, 'bold')
        
        # Banner de título con color de fondo
        banner_frame = tk.Frame(main_frame, bg='#673AB7', relief=tk.RAISED, bd=3)
        banner_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 25))
        
        titulo = tk.Label(banner_frame, text="⚕ GENERADOR DE CONSENTIMIENTOS - SEPAS", 
                          font=fuente_titulo, bg='#673AB7', fg='white', pady=15)
        titulo.pack(fill=tk.X)
        
        # Sección 1: Archivo Excel
        tk.Label(main_frame, text="➊ Archivo Excel con datos:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').grid(row=1, column=0, columnspan=3, 
                                                   sticky=tk.W, pady=(10, 5))
        
        entry_excel = ttk.Entry(main_frame, textvariable=self.archivo_excel, 
                 width=60, font=fuente_normal)
        entry_excel.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                               padx=(0, 10), pady=5)
        btn_excel = tk.Button(main_frame, text="📁 Seleccionar Excel", 
                  command=self.seleccionar_excel, font=fuente_boton,
                  bg='#4CAF50', fg='white', activebackground='#45a049',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_excel.grid(row=2, column=2, sticky=tk.W, padx=5)
        
        # Sección 2: Plantilla Word
        tk.Label(main_frame, text="➋ Plantilla Word (consentimiento):", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').grid(row=3, column=0, columnspan=3, 
                                                   sticky=tk.W, pady=(20, 5))
        
        entry_word = ttk.Entry(main_frame, textvariable=self.plantilla_word, 
                 width=60, font=fuente_normal)
        entry_word.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                               padx=(0, 10), pady=5)
        btn_word = tk.Button(main_frame, text="📄 Seleccionar Plantilla", 
                  command=self.seleccionar_plantilla, font=fuente_boton,
                  bg='#2196F3', fg='white', activebackground='#0b7dda',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_word.grid(row=4, column=2, sticky=tk.W, padx=5)
        
        # Sección 3: Carpeta de salida
        tk.Label(main_frame, text="➌ Carpeta para guardar consentimientos:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').grid(row=5, column=0, columnspan=3, 
                                                   sticky=tk.W, pady=(20, 5))
        
        entry_salida = ttk.Entry(main_frame, textvariable=self.carpeta_salida, 
                 width=60, font=fuente_normal)
        entry_salida.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E), 
                               padx=(0, 10), pady=5)
        btn_carpeta = tk.Button(main_frame, text="📂 Seleccionar Carpeta", 
                  command=self.seleccionar_carpeta, font=fuente_boton,
                  bg='#FF9800', fg='white', activebackground='#e68900',
                  cursor='hand2', relief=tk.RAISED, bd=3, padx=20, pady=10)
        btn_carpeta.grid(row=6, column=2, sticky=tk.W, padx=5)
        
        # Botón principal
        btn_frame = tk.Frame(main_frame, bg='#f5f5f5')
        btn_frame.grid(row=7, column=0, columnspan=3, pady=30)
        
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
        btn_emails = tk.Button(btn_frame, text="📧 PREPARAR\nEMAILS", 
                              command=self.preparar_emails,
                              font=fuente_boton,
                              bg='#00BCD4', fg='white',
                              activebackground='#0097A7',
                              cursor='hand2', relief=tk.RAISED, bd=4,
                              padx=30, pady=20, width=20, height=4)
        btn_emails.pack(side=tk.LEFT, padx=10)
        
        # Botón para abrir carpeta
        btn_abrir_carpeta = tk.Button(btn_frame, text="📁 ABRIR\nCARPETA\nDE PDFs", 
                              command=self.abrir_carpeta_salida,
                              font=fuente_boton,
                              bg='#4CAF50', fg='white',
                              activebackground='#45a049',
                              cursor='hand2', relief=tk.RAISED, bd=4,
                              padx=30, pady=20, width=20, height=4)
        btn_abrir_carpeta.pack(side=tk.LEFT, padx=10)
        
        # Separador
        ttk.Separator(main_frame, orient='horizontal').grid(row=8, column=0, 
                                                            columnspan=3, 
                                                            sticky=(tk.W, tk.E), 
                                                            pady=10)
        
        # Área de log
        tk.Label(main_frame, text="► Registro de actividad:", 
                 font=fuente_seccion, bg='#f5f5f5', fg='#333').grid(row=9, column=0, columnspan=3, 
                                                   sticky=tk.W, pady=(10, 5))
        
        self.log_text = scrolledtext.ScrolledText(main_frame, height=15, width=80, 
                                                  wrap=tk.WORD, font=fuente_normal)
        self.log_text.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), 
                          pady=(0, 10))
        
        # Configurar expansión del log
        main_frame.rowconfigure(10, weight=1)
        
        # Botón limpiar log
        btn_limpiar = tk.Button(main_frame, text="🗑️ Limpiar Log", 
                  command=self.limpiar_log, font=fuente_boton,
                  bg='#607D8B', fg='white', activebackground='#455A64',
                  cursor='hand2', relief=tk.RAISED, bd=2, padx=15, pady=8)
        btn_limpiar.grid(row=11, column=0, columnspan=3, pady=10)
        
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
        # Abrir diálogo en la carpeta docs del programa
        carpeta_docs = Path(__file__).parent / "docs"
        carpeta_inicial = str(carpeta_docs) if carpeta_docs.exists() else str(Path.cwd())
        
        archivo = filedialog.askopenfilename(
            title="Seleccionar plantilla Word",
            initialdir=carpeta_inicial,
            filetypes=[("Documentos Word", "*.docx"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.plantilla_word.set(archivo)
            self.agregar_log(f"✓ Plantilla seleccionada: {Path(archivo).name}")
            
    def seleccionar_carpeta(self):
        """Abre diálogo para seleccionar carpeta de salida"""
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if carpeta:
            self.carpeta_salida.set(carpeta)
            self.agregar_log(f"✓ Carpeta de salida: {carpeta}")
            
    def agregar_log(self, mensaje):
        """Agrega un mensaje al área de log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {mensaje}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def limpiar_log(self):
        """Limpia el área de log"""
        self.log_text.delete(1.0, tk.END)
        
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
        
    def generar_consentimientos(self):
        """Genera los consentimientos desde Excel y plantilla Word"""
        if not self.validar_archivos():
            return
        
        if not self.mapeo_campos:
            messagebox.showwarning("Advertencia",
                                 "No has configurado el mapeo de campos.\n"
                                 "Selecciona el archivo Excel para configurar el mapeo.")
            return
            
        try:
            from modulos.lector_excel import LectorExcel
            from modulos.generador_word import GeneradorConsentimientos
            
            self.agregar_log("=" * 80)
            self.agregar_log("Iniciando proceso de generación de consentimientos en PDF...")
            
            # Crear carpeta de salida si no existe
            carpeta_salida = Path(self.carpeta_salida.get())
            carpeta_salida.mkdir(parents=True, exist_ok=True)
            
            # Leer datos del Excel
            self.agregar_log(f"Leyendo datos de: {Path(self.archivo_excel.get()).name}")
            lector = LectorExcel(self.archivo_excel.get())
            datos = lector.leer_datos()
            
            self.agregar_log(f"✓ Se encontraron {len(datos)} registros en el Excel")
            self.agregar_log("⚠️ MODO PRUEBA: Solo se generará el PDF del PRIMER USUARIO")
            
            # Tomar solo el primer registro
            datos = datos[:1] if datos else []
            
            if not datos:
                self.agregar_log("✗ Error: No hay datos en el Excel")
                messagebox.showerror("Error", "No hay datos en el Excel")
                return
            
            # Aplicar mapeo a los datos
            self.agregar_log(f"Aplicando mapeo de campos: {len(self.mapeo_campos)} campos")
            if self.valores_fijos:
                self.agregar_log(f"Aplicando valores fijos: {len(self.valores_fijos)} campos")
            
            datos_mapeados = []
            for registro in datos:
                registro_nuevo = {}
                
                # Primero aplicar mapeo desde Excel
                for campo_plantilla, columna_excel in self.mapeo_campos.items():
                    if columna_excel in registro:
                        valor = registro[columna_excel]
                        registro_nuevo[campo_plantilla] = valor
                
                # Luego aplicar valores fijos (solo si no se mapearon desde Excel)
                for campo, valor_fijo in self.valores_fijos.items():
                    if campo not in registro_nuevo:
                        registro_nuevo[campo] = valor_fijo
                
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
            
            # Generar consentimientos
            generador = GeneradorConsentimientos(self.plantilla_word.get(), 
                                                str(carpeta_salida))
            
            consentimientos_generados = []
            errores = 0
            
            for i, registro in enumerate(datos_mapeados, 1):
                try:
                    # Mostrar qué datos se están usando
                    self.agregar_log(f"  Datos para registro {i}: {registro}")
                    archivo_generado = generador.generar_consentimiento(registro)
                    consentimientos_generados.append(archivo_generado)
                    self.agregar_log(f"  [{i}/{len(datos_mapeados)}] ✓ Generado: {Path(archivo_generado).name}")
                except Exception as e:
                    errores += 1
                    self.agregar_log(f"  [{i}/{len(datos_mapeados)}] ✗ Error: {str(e)}")
            
            self.agregar_log("=" * 80)
            self.agregar_log(f"RESUMEN: {len(consentimientos_generados)} consentimientos generados correctamente")
            if errores > 0:
                self.agregar_log(f"ADVERTENCIA: {errores} registros con errores")
            
            messagebox.showinfo("Completado", 
                              f"Se generaron {len(consentimientos_generados)} consentimientos en PDF\n"
                              f"Carpeta: {carpeta_salida}")
            
        except ImportError as e:
            self.agregar_log(f"✗ Error: Módulos no encontrados - {str(e)}")
            messagebox.showerror("Error", 
                               "Faltan módulos necesarios. Asegúrese de instalar:\n"
                               "pip install openpyxl python-docx")
        except Exception as e:
            self.agregar_log(f"✗ Error: {str(e)}")
            messagebox.showerror("Error", f"Error al generar consentimientos:\n{str(e)}")
            
    def preparar_emails(self):
        """Prepara los emails con los consentimientos adjuntos"""
        carpeta_salida = Path(self.carpeta_salida.get())
        
        if not carpeta_salida.exists():
            messagebox.showerror("Error", "No hay consentimientos generados")
            return
            
        try:
            from modulos.preparador_emails import PreparadorEmails
            
            self.agregar_log("=" * 80)
            self.agregar_log("Preparando emails...")
            
            preparador = PreparadorEmails(str(carpeta_salida))
            resultado = preparador.crear_borradores_email(self.archivo_excel.get())
            
            self.agregar_log(f"✓ Se prepararon {len(resultado)} borradores de email")
            self.agregar_log(f"✓ Archivo de emails generado en: {carpeta_salida}")
            
            messagebox.showinfo("Completado", 
                              f"Se prepararon {len(resultado)} emails\n"
                              "Revise el archivo 'emails_para_enviar.txt' en la carpeta de salida")
            
        except ImportError:
            self.agregar_log("✗ Error: Módulo preparador_emails no encontrado")
            messagebox.showerror("Error", "Módulo preparador_emails no encontrado")
        except Exception as e:
            self.agregar_log(f"✗ Error: {str(e)}")
            messagebox.showerror("Error", f"Error al preparar emails:\n{str(e)}")
    
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
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Crear comboboxes para cada campo necesario
            self.comboboxes_mapeo = {}
            self.textboxes_fijos = {}  # Para valores fijos opcionales
            columnas_con_vacio = ["(No mapear - dejar vacío)"] + columnas_excel
            
            # Campos que pueden tener valores fijos
            campos_con_valor_fijo = ['codigo_postal', 'poblacion', 'provincia', 'pais_deudor']
            
            for i, (campo_interno, descripcion) in enumerate(campos_necesarios):
                # Frame para cada fila
                fila_frame = tk.Frame(scrollable_frame, bg='white', relief=tk.RAISED, bd=1)
                fila_frame.pack(fill=tk.X, pady=5, padx=5)
                
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
                    'swift_bic': ['swift', 'bic', 'banco', 'bank'],
                    'iban': ['iban', 'cuenta', 'account', 'numero cuenta'],
                    'identificador_acreedor': ['identificador', 'id acreedor', 'creditor'],
                    'nombre_acreedor': ['acreedor', 'empresa', 'organizacion', 'creditor'],
                    'direccion_acreedor': ['direccion acreedor', 'domicilio social'],
                    'firma': ['firma', 'signature', 'signatura']
                }
                
                # Buscar coincidencias usando las palabras clave
                mejor_coincidencia = None
                mejor_puntuacion = 0
                
                for columna in columnas_excel:
                    columna_lower = columna.lower()
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


if __name__ 