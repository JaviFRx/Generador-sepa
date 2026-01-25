# Sistema de Generación de Consentimientos - SEPAS

Sistema automatizado para generar consentimientos personalizados desde plantillas Word utilizando datos de Excel y prepararlos para envío por correo electrónico.

## 📋 Características

- **Lectura de Excel**: Lee datos de participantes desde archivos Excel (.xlsx, .xls)
- **Generación automática**: Crea documentos Word personalizados desde una plantilla
- **Preparación de emails**: Genera borradores de correos electrónicos listos para enviar
- **Interfaz gráfica**: Aplicación de escritorio fácil de usar
- **Registro de actividades**: Log detallado de todas las operaciones

## 🚀 Instalación

### Requisitos previos
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Pasos de instalación

1. **Clonar o descargar el proyecto**

2. **Instalar dependencias**
   ```bash
   pip install -r requirements.txt
   ```

   O instalar manualmente:
   ```bash
   pip install openpyxl python-docx
   ```

## 📖 Uso

### 1. Preparar archivos

#### Archivo Excel
Debe contener una hoja con:
- **Primera fila**: Encabezados (nombre, apellido, email, dni, fecha, etc.)
- **Siguientes filas**: Datos de cada participante

Ejemplo:
```
| nombre | apellido | email              | dni      | fecha      |
|--------|----------|--------------------|----------|------------|
| Juan   | Pérez    | juan@email.com     | 12345678 | 25/01/2026 |
| María  | García   | maria@email.com    | 87654321 | 25/01/2026 |
```

#### Plantilla Word
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
