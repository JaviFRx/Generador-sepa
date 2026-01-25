# Guía de Uso Rápido

## Inicio Rápido - 3 Pasos

### 1️⃣ Instalar Dependencias
```bash
pip install -r requirements.txt
```

### 2️⃣ Ejecutar la Aplicación
```bash
python app_consentimientos.py
```

### 3️⃣ Usar la Aplicación
1. Selecciona tu archivo Excel
2. Selecciona tu plantilla Word
3. Click en "Generar Consentimientos"
4. Click en "Preparar Emails"

---

## Archivos de Ejemplo Incluidos

En la carpeta `ejemplos/` encontrarás:
- **datos_ejemplo.xlsx**: Excel con datos de muestra
- **plantilla_consentimiento.docx**: Plantilla Word de ejemplo

Puedes usar estos archivos para probar el sistema.

---

## Formato de la Plantilla Word

### Marcadores Disponibles
Usa estos marcadores en tu documento Word (con doble llave):

- `{{nombre}}` - Nombre de la persona
- `{{apellido}}` - Apellido de la persona
- `{{email}}` - Correo electrónico
- `{{dni}}` - Número de documento
- `{{fecha}}` - Fecha

### Ejemplo de Uso en Word:
```
CONSENTIMIENTO INFORMADO

Yo, {{nombre}} {{apellido}}, con DNI {{dni}}, 
declaro que autorizo...

Fecha: {{fecha}}
Contacto: {{email}}

Firma: _________________
```

---

## Formato del Excel

### Estructura Requerida:

| nombre | apellido | email | dni | fecha |
|--------|----------|-------|-----|-------|
| Juan | Pérez | juan@email.com | 12345678 | 25/01/2026 |
| María | García | maria@email.com | 87654321 | 25/01/2026 |

**Importante:**
- Primera fila = encabezados
- Siguientes filas = datos
- Los nombres de columnas deben coincidir con los marcadores

---

## Resultados

Después de ejecutar el proceso, obtendrás:

### 📁 Consentimientos Generados
- Un archivo .docx por cada persona
- Nombrados automáticamente: `Apellido_Nombre_DNI.docx`
- Guardados en la carpeta de salida

### 📧 Archivos de Email
- `emails_para_enviar.txt` - Borradores completos de email
- `emails_lista.csv` - Lista para importar a cliente de correo

---

## Solución Rápida de Problemas

### ❌ "Módulos no encontrados"
```bash
pip install openpyxl python-docx
```

### ❌ "No se puede leer el Excel"
- Cierra el archivo Excel
- Verifica que sea .xlsx o .xls
- Verifica que tenga encabezados en fila 1

### ❌ "Los marcadores no se reemplazan"
- Usa doble llave: `{{campo}}`
- Sin espacios: ❌ `{{ campo }}` ✅ `{{campo}}`
- Nombres deben coincidir con columnas Excel

---

## Contacto y Soporte

Para preguntas o problemas, contacta al administrador del sistema.

---

**SEPAS © 2026**
