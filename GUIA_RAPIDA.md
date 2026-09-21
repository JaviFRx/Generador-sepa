# Guía de Uso Rápido

## Inicio Rápido - 5 Pasos

### 1️⃣ Instalar Dependencias
```bash
pip install -r requirements.txt
```

### 2️⃣ Ejecutar la Aplicación
- Opción A (sin instalar nada): ejecuta el archivo `dist/Generador_Consentimientos_SEPAS.exe`
- Opción B (con Python):
```bash
python app_consentimientos.py
```

### 3️⃣ Configurar Gmail (Primera vez)
En la sección **"📧 Envío de Emails"**:
- Introduce tu correo Gmail (ej: tu_email@gmail.com)
- Introduce tu **Contraseña de Aplicación** (16 caracteres especiales)
  - [Ver: Cómo generar Contraseña de Aplicación](#-cómo-generar-contraseña-de-aplicación-google)
- Click en **"💾 Guardar Configuración"**

### 4️⃣ Generar Consentimientos
1. Selecciona tu archivo Excel con datos
2. Selecciona tu plantilla Word
3. Mapea los campos (asegúrate que coincidan con Excel)
4. Click en **"📝 Generar Consentimientos"**

### 5️⃣ Crear borradores y revisarlos en Gmail
1. Pulsa **Configurar correo** y personaliza el asunto y cuerpo
2. Revisa destinatarios y adjuntos en `emails_lista.csv`
3. Con tu cuenta y contraseña de aplicación configuradas, pulsa **Crear borradores (Gmail)**
4. Abre **Gmail → Borradores** con esa misma cuenta
5. Comprueba cada destinatario y su PDF en Gmail
6. Vuelve a la app y pulsa **Enviar borradores revisados (Gmail)**
7. Revisa la lista y confirma **Enviar los N borradores revisados**

Crear borradores no envía correos: el envío requiere el botón y la confirmación
posteriores. Si regeneras los PDFs, revisa también los borradores
antiguos de Gmail: la limpieza de PDFs solo afecta a los archivos locales.

Conserva `historial_borradores.json`: evita repetir los envíos confirmados.
Si se avisa de un envío sin confirmar, comprueba **Enviados** en Gmail antes de reintentar.

---

## Archivos de Ejemplo Incluidos

En la carpeta `ejemplos/` encontrarás:
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
- `emails_para_enviar.txt` - Borradores completos de email (revisa antes de enviar)
- `emails_lista.csv` - Lista con información de destinatarios

### 📧 Gmail - Emails Enviados
- Se enviarán automáticamente desde tu cuenta Gmail
- Los PDFs se adjuntarán a cada email
- Las credenciales se guardan encriptadas localmente

---

## 🔑 Cómo Generar Contraseña de Aplicación Google

**IMPORTANTE**: No puedes usar tu contraseña normal de Gmail. Debes generar una **Contraseña de Aplicación**.

### Pasos:

1. **Ve a tu cuenta Google**: https://myaccount.google.com
2. **Seguridad** (lado izquierdo)
3. **Autenticación de dos factores** 
   - Si no está activada, actívala ahora (requiere número de teléfono)
4. **Contraseñas de aplicación** (aparece abajo después de activar 2FA)
5. Selecciona:
   - Dispositivo: **Correo**
   - Aplicación: **Windows**
6. Google te generará una contraseña de **16 caracteres**
7. Copia esa contraseña
8. Pégala en el campo de "Contraseña" en la aplicación

✅ **Listo**: Usa esa contraseña de 16 caracteres en la aplicación

---

## Solución Rápida de Problemas

### ❌ "Módulos no encontrados"
```bash
pip install -r requirements.txt
```

### ❌ "No se puede leer el Excel"
- Cierra el archivo Excel
- Verifica que sea .xlsx o .xls
- Verifica que tenga encabezados en fila 1

### ❌ "Los marcadores no se reemplazan"
- Usa doble llave: `{{campo}}`
- Sin espacios: ❌ `{{ campo }}` ✅ `{{campo}}`
- Nombres deben coincidir con columnas Excel

### ❌ Gmail - "Credenciales inválidas"
- Verifica que uses **Contraseña de Aplicación**, no contraseña normal
- Verifica que el email incluya `@gmail.com`
- Si cambias de contraseña, actualiza en "💾 Guardar Configuración"

### ❌ "PDF no encontrado al enviar"
- Verifica que los PDFs se generaron correctamente (paso 4)
- Revisa los nombres de columnas en Excel y en el mapeo de campos
- Los PDFs deben estar en la misma carpeta que los consentimientos

---

## Contacto y Soporte

Para preguntas o problemas, contacta al administrador del sistema.

---

Software propiedad de Javier Fernández Ramos. Uso sujeto a autorización del autor.
