# Generador de Consentimientos SEPA 🏦

Sistema automatizado para generar documentos de domiciliación bancaria SEPA (Single Euro Payments Area) en formato PDF con campos editables, a partir de datos de Excel.

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## ✨ Características Principales

- 📊 **Lectura inteligente de Excel**: Importa datos desde archivos .xlsx o .xls
- 🎯 **Mapeo interactivo de campos**: Sistema visual para asignar columnas del Excel a campos del formulario
- 🤖 **Auto-detección inteligente**: Reconoce automáticamente columnas por palabras clave (español/catalán)
- 🌍 **Auto-completado geográfico**: Busca automáticamente código postal ↔ población/provincia
- 📝 **PDF con campos editables**: Genera PDFs con campos de fecha, localidad y firma editables
- 📧 **Preparación de emails**: Crea borradores de correos listos para enviar
- 🎨 **Interfaz moderna**: GUI intuitiva con diseño colorido y responsive
- 📁 **Gestión de archivos**: Apertura directa de carpeta de salida

## 🖼️ Capturas de Pantalla

### Pantalla Principal
- Interfaz con botones grandes y coloridos
- Generación de PDFs, preparación de emails y apertura de carpeta de salida

### Ventana de Mapeo Inteligente
- Sistema guiado campo por campo
- Auto-detección de columnas
- Valores fijos para campos comunes
- Auto-completado geográfico

## 📋 Requisitos

- **Python 3.10 o superior**
- **Microsoft Word** (para conversión a PDF)
- **Conexión a internet** (opcional, para auto-completado geográfico)

## 🚀 Instalación

### 1️⃣ Clonar el repositorio

```bash
git clone https://github.com/JaviFRx/Generador-sepa.git
cd Generador-sepa
```

### 2️⃣ Crear entorno virtual

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Linux/Mac
python3 -m venv .venv
source .venv/bin/activate
```

### 3️⃣ Instalar dependencias

```bash
pip install -r requirements.txt
```

## 📦 Dependencias

```
openpyxl==3.1.2       # Lectura de archivos Excel
python-docx==1.1.0    # Manipulación de documentos Word
docx2pdf==0.1.8       # Conversión de Word a PDF
pypdf==6.6.1          # Creación de campos editables en PDF
```

## 🎯 Uso Básico

### 1. Ejecutar la aplicación

```bash
python app_consentimientos.py
```

### 2. Preparar tu plantilla Word

Añade marcadores en tu plantilla con formato `{{nombre_campo}}`:

```
Referencia de la orden de domiciliación: {{referencia_orden}}
Nombre del deudor: {{nombre_deudor}}
Dirección: {{direccion_deudor}}
Código postal: {{codigo_postal}}
Población: {{poblacion}}
Provincia: {{provincia}}
País: {{pais_deudor}}
IBAN: {{iban}}
Swift BIC: {{swift_bic}}

En {{localidad_firma}}, a {{fecha}}

Firma del deudor: {{firma}}
```

### 3. Preparar tu Excel

Tu Excel debe tener columnas como:
- Nombre del alumno / Nom de l'alumne
- Apellidos / Cognoms
- Nombre del titular / Titular compte bancari
- Dirección / Adreça
- Código postal / Codi postal
- Población / Població
- IBAN
- Swift BIC
- etc.

### 4. Proceso de generación

1. **Seleccionar Excel** → Se abre ventana de mapeo automáticamente
2. **Configurar mapeo**:
   - Asigna cada campo al nombre de columna del Excel
   - El sistema auto-detecta las columnas más probables
   - Puedes usar valores fijos (ej: mismo código postal para todos)
3. **Auto-completado geográfico**:
   - Escribe código postal → rellena población, provincia, país
   - O escribe población → rellena código postal y provincia
4. **Generar PDFs**: Crea documentos con campos editables

## 🔖 Campos SEPA Soportados

### Datos del Alumno
- `{{referencia_orden}}` - Auto-generado: Nombre_Apellido_Robotica

### Datos del Deudor
- `{{nombre_deudor}}` - Nombre del titular de la cuenta
- `{{direccion_deudor}}` - Dirección completa
- `{{codigo_postal}}` - Código postal
- `{{poblacion}}` - Población/Ciudad
- `{{provincia}}` - Provincia
- `{{pais_deudor}}` - País

### Datos Bancarios
- `{{iban}}` - Número de cuenta IBAN
- `{{swift_bic}}` - Código Swift BIC del banco

### Datos del Acreedor
- `{{identificador_acreedor}}` - Identificador del acreedor
- `{{nombre_acreedor}}` - Nombre del acreedor
- `{{direccion_acreedor}}` - Dirección del acreedor

### Campos Auto-generados
- `{{fecha}}` - Fecha actual (formato: DD/MM/AAAA)
- `{{localidad_firma}}` - Copia de la población del deudor

### Campos Editables en PDF
- **Fecha** - Campo editable para firma
- **Localidad** - Campo editable para lugar de firma
- **Firma** - Campo editable para firma del deudor

## 🧠 Auto-detección Inteligente

El sistema detecta automáticamente columnas usando palabras clave:

| Campo | Palabras clave detectadas |
|-------|--------------------------|
| Nombre alumno | nombre, nom, name, alumne, alumno |
| Apellido | apellido, cognom, surname, cognoms |
| IBAN | iban, cuenta, account, bancaria, compte |
| Código postal | cp, codigo postal, codi postal, zip |
| Población | poblacion, població, ciudad, localidad |
| Provincia | provincia, prov, province |
| etc. | ... |

## 📁 Estructura del Proyecto

```
Generador-sepa/
├── app_consentimientos.py              # Aplicación principal
├── modulos/
│   ├── __init__.py
│   ├── lector_excel.py                 # Lectura de Excel
│   ├── generador_word.py               # Generación Word/PDF con campos editables
│   └── preparador_emails.py            # Preparación de emails
├── docs/
│   └── plantilla_domiciliacion_sepa.docx  # Plantilla SEPA
├── ejemplos/
│   └── datos_ejemplo.xlsx              # Excel de ejemplo
├── requirements.txt                    # Dependencias
├── .gitignore                         # Archivos ignorados
└── README.md                          # Este archivo
```

## ⚙️ Configuración Avanzada

### Personalizar campos editables

Edita `modulos/generador_word.py`, método `_anadir_campos_editables()`:

```python
campos = [
    {'nombre': 'fecha', 'x': 150, 'y': page_height - 680, 'width': 120, 'height': 20},
    {'nombre': 'localidad', 'x': 350, 'y': page_height - 680, 'width': 200, 'height': 20},
    {'nombre': 'firma', 'x': 150, 'y': page_height - 720, 'width': 400, 'height': 60}
]
```

Ajusta valores de `x`, `y`, `width` y `height` según tu plantilla.

## 🐛 Solución de Problemas

### Error: "Microsoft Word no está instalado"
- **Solución**: Instala Microsoft Word o usa una alternativa compatible con docx2pdf

### Los campos no se auto-detectan
- **Solución**: Renombra las columnas del Excel con nombres más descriptivos

### El auto-completado geográfico no funciona
- **Solución**: Verifica tu conexión a internet. Usa valores fijos si no tienes conexión

### Los campos editables no aparecen en el PDF
- **Solución**: Abre el PDF con Adobe Reader o un lector compatible con formularios PDF

## 🤝 Contribuciones

Las contribuciones son bienvenidas:

1. Fork el proyecto
2. Crea una rama (`git checkout -b feature/NuevaCaracteristica`)
3. Commit tus cambios (`git commit -m 'Añadir nueva característica'`)
4. Push a la rama (`git push origin feature/NuevaCaracteristica`)
5. Abre un Pull Request

## 📜 Licencia

Este proyecto está bajo la Licencia MIT. Ver archivo `LICENSE` para más detalles.

## 🙏 Agradecimientos

- **APIs utilizadas**: 
  - [Zippopotam.us](http://zippopotam.us) - Búsqueda de códigos postales
  - [OpenStreetMap Nominatim](https://nominatim.openstreetmap.org) - Geocodificación
- **Librerías Python**: openpyxl, python-docx, docx2pdf, pypdf

## 👤 Autor

**JaviFRx**

- GitHub: [@JaviFRx](https://github.com/JaviFRx)
- Proyecto: [Generador-sepa](https://github.com/JaviFRx/Generador-sepa)

## 📝 Changelog

### v1.0.0 (2026-01-25)
- ✨ Mapeo inteligente de campos con auto-detección
- 🌍 Auto-completado geográfico (código postal ↔ población)
- 📝 PDFs con campos editables (fecha, localidad, firma)
- 🎨 Interfaz gráfica moderna y responsive
- 🔄 Scroll con rueda del ratón en ventana de mapeo
- 📊 Soporte multi-idioma (español/catalán)

---

⭐ Si este proyecto te ha sido útil, considera darle una estrella en GitHub!
Cree un documento Word (.docx) con marcadores usando el formato `{{campo}}`:
- `{{nombre}}` - Se reemplazará por el nombre
- `{{apellido}}` - Se reemplazará por el apellido
- `{{email}}` - Se reemplazará por el email
- `{{dni}}` - Se reemplazará por el DNI
- `{{fecha}}` - Se reemplazará por la fecha

**Ejemplo de plantilla:**
```
CONSENTIMIENTO INFORMADO

Yo, {{nombre}} {{apellido}}, con DNI {{dni}}, declaro que:

1. He sido informado/a sobre...
2. Autorizo...

Fecha: {{fecha}}

Firma: _________________
```

### 2. Ejecutar la aplicación

```bash
python app_consentimientos.py
```

### 3. Usar la interfaz

1. **Seleccionar Excel**: Click en "Seleccionar Excel" y elegir su archivo
2. **Seleccionar Plantilla**: Click en "Seleccionar Plantilla" y elegir su documento Word
3. **Seleccionar Carpeta de Salida**: (Opcional) Elegir dónde guardar los archivos generados
4. **Generar Consentimientos**: Click en "Leer Excel y Generar Consentimientos"
5. **Preparar Emails**: Click en "Preparar Emails" para crear borradores de correo

## 📁 Estructura del Proyecto

```
Sepas/
│
├── app_consentimientos.py          # Aplicación principal
├── modulos/
│   ├── __init__.py
│   ├── lector_excel.py             # Lectura de archivos Excel
│   ├── generador_word.py           # Generación de documentos Word
│   └── preparador_emails.py        # Preparación de emails
│
├── ejemplos/
│   ├── plantilla_consentimiento.docx   # Plantilla de ejemplo
│   └── datos_ejemplo.xlsx               # Datos de ejemplo
│
├── requirements.txt                # Dependencias del proyecto
└── README.md                      # Este archivo
```

## 📤 Archivos Generados

Al ejecutar el programa, se crearán:

1. **Consentimientos individuales**: 
   - Formato: `Apellido_Nombre_DNI.docx`
   - Ubicación: Carpeta de salida seleccionada

2. **Lista de emails** (`emails_para_enviar.txt`):
   - Contiene todos los borradores de email
   - Incluye destinatario, asunto y cuerpo del mensaje

3. **CSV de emails** (`emails_lista.csv`):
   - Formato CSV para importar a clientes de correo
   - Útil para envíos masivos

## 🔧 Personalización

### Modificar el cuerpo del email
Edite el método `_generar_cuerpo_email` en `modulos/preparador_emails.py`:

```python
def _generar_cuerpo_email(self, nombre: str) -> str:
    return f"""Su mensaje personalizado aquí...
    
Estimado/a {nombre},
...
"""
```

### Cambiar formato de nombres de archivo
Edite el método `_generar_nombre_archivo` en `modulos/generador_word.py`

## ⚠️ Solución de Problemas

### Error: "Módulos no encontrados"
**Solución**: Instalar dependencias
```bash
pip install openpyxl python-docx
```

### Error: "El archivo Excel no se puede leer"
**Soluciones**:
- Verificar que el archivo sea .xlsx o .xls
- Cerrar el archivo si está abierto en Excel
- Verificar que el archivo tenga encabezados en la primera fila

### Los marcadores no se reemplazan
**Soluciones**:
- Verificar que los marcadores usen `{{}}` dobles llaves
- Asegurarse que los nombres coincidan con los encabezados del Excel
- Revisar que no haya espacios dentro de los marcadores

### No se encuentran los archivos de consentimiento al preparar emails
**Solución**: Verificar que los archivos se hayan generado correctamente antes de preparar emails

## 📝 Notas Importantes

- Los nombres de columnas en Excel no distinguen mayúsculas/minúsculas
- Los consentimientos se sobrescriben si ya existe un archivo con el mismo nombre
- Se recomienda revisar los borradores antes de enviar emails
- Los archivos generados usan formato .docx (Word 2007+)

## 🤝 Soporte

Para problemas o preguntas, contacte al administrador del sistema.

## 📄 Licencia

Este software es propiedad de SEPAS. Todos los derechos reservados.
