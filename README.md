# Generador de Consentimientos SEPA

Aplicación para Windows que genera PDFs de domiciliación SEPA desde un Excel y una
plantilla Word. Permite revisar cada PDF y enviarlo por Gmail al destinatario
asociado, con comprobaciones antes del envío y de la copia en Enviados.

- [Guía rápida](GUIA_RAPIDA.md): uso diario y resolución de errores.
- [Ejecutable](INSTRUCCIONES_EXE.md): compilación y actualización.
- [Documentación técnica](DOCUMENTACION_TECNICA.md): validaciones, historial y pruebas.

## Requisitos e inicio

- Windows y Microsoft Word instalado para convertir los documentos a PDF.
- Excel `.xlsx`, con encabezados en la primera fila. Convierte los `.xls` antiguos
  a `.xlsx`. Se utiliza la hoja activa del libro.
- Plantilla `.docx` con marcadores `{{campo}}`.
- Cuenta Gmail y contraseña de aplicación para enviar; conexión a internet para
  Gmail y para las consultas geográficas opcionales.

Ejecuta `dist/Generador_Consentimientos_SEPAS.exe`. La entrega actual está en
`E:/Sepas/dist/Generador_Consentimientos_SEPAS.exe`. Incluye Python y dependencias,
pero no Microsoft Word.

Para ejecutar desde el código con Python 3.10 o posterior, en PowerShell:

```powershell
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements.txt
.venv/Scripts/python.exe app_consentimientos.py
```

Las versiones de las dependencias están en [requirements.txt](requirements.txt).

## Flujo de trabajo

1. Selecciona el Excel, la plantilla Word y la carpeta de salida.
2. Revisa el mapeo de columnas y campos. Puedes usar valores fijos y completar
   población, provincia y código postal.
3. Pulsa **Generar consentimientos en PDF**. Una modal muestra procesados,
   pendientes, errores, tiempo transcurrido y estimación mientras Word trabaja
   en segundo plano.
4. Pulsa **Configurar correo**. Personaliza asunto y texto; `{nombre}` inserta el
   nombre detectado. Se generan listas locales para revisión.
5. Pulsa **Enviar correos (Gmail)**. La app comprueba todo el lote y muestra
   destinatario, asunto y PDF de cada correo pendiente.
6. **Haz clic en una fila para abrir su PDF** en el visor predeterminado.
   También puedes seleccionar una fila y pulsar Intro. Abrir el PDF no envía nada.
7. Confirma **Enviar los N correos** o cancela.
8. Revisa el resumen y `informe_envios.csv` en la carpeta de salida.

## Excel y destinatarios

Se reconocen encabezados que contengan `email`, `mail`, `correo` o `correu`, sin
distinguir mayúsculas. Por ejemplo, `Direcció correu electrònic (minúscula)`.
Cada registro necesita un único correo válido. Varias columnas con el mismo
correo se admiten; direcciones distintas en una misma fila bloquean el lote.
No se elige automáticamente una de las direcciones.

Los errores identifican la hoja, la fila real del Excel, las columnas y el motivo:
campo vacío, formato inválido, varias direcciones o encabezado no reconocido.
Se conserva la numeración aunque haya filas vacías. El detalle se guarda en
`errores_correos.txt` cuando es posible. Si corriges datos del Excel, regenera
los PDFs para restablecer la asociación del lote.

## Protección de los adjuntos

`registro_consentimientos.json` asocia cada fila del Excel con su PDF mediante
huellas SHA-256. Antes de preparar o enviar se comprueba que:

- El Excel conserva sus registros y su orden.
- La generación terminó y están todos los PDFs.
- Cada archivo está en la carpeta de salida y conserva su contenido.
- No se reutiliza la misma ruta ni un PDF de contenido idéntico entre registros.
- El destinatario y el adjunto corresponden a su fila original. No se admiten
  CC/CCO ni adjuntos adicionales en el correo confirmado.

Al confirmar se repiten las comprobaciones. Al abrir un PDF desde la lista también
se verifica que mantiene el contenido revisado.

Una nueva generación válida **borra los PDFs anteriores y las listas locales**
`emails_lista.csv` y `emails_para_enviar.txt`. Conserva historiales, informes,
otros archivos y subcarpetas. Si no puede borrar un PDF, se detiene.
Los nombres nuevos incorporan un identificador único.

## Interpretar el resultado

| Resultado | Qué acredita |
| --- | --- |
| Aceptados por Gmail | Gmail respondió con aceptación SMTP del mensaje. |
| Comprobados en Enviados | Se encontró una copia con identificador, destinatario, remitente, asunto y PDF esperados. |
| Ya enviados anteriormente | El historial ya registra ese correo como enviado; se omite. No es una comprobación nueva. |
| Pendientes de enviar | No se registró aceptación en esta operación. Si hubo error de conexión, revisa el aviso de posible envío incierto. |

**Aceptación y presencia en Enviados no acreditan entrega ni lectura por el
destinatario.** Si la copia no se puede comprobar tras tres consultas, se detiene
el lote y se avisa. El correo aceptado no se repite automáticamente.

El estado se registra antes de contactar con SMTP. Una pérdida de respuesta deja
el envío sin confirmar y bloquea su repetición. No borres ni edites el historial
para forzar un reenvío: primero comprueba qué ocurrió.

La protección usa el historial de la cuenta y los archivos del lote. Regenerar
crea archivos nuevos; borrar el historial o mover el lote sin él puede perder esa
protección. Revisa los envíos anteriores antes de enviar un lote regenerado.

## Borradores de versiones anteriores

La interfaz actual utiliza **envío directo**, con un identificador nuevo por correo.
No crea ni elimina borradores en Gmail. No envíes manualmente los borradores
antiguos si la aplicación ya envió ese correo.

Se mantiene `historial_borradores.json` para conservar los estados anteriores.
Un registro antiguo `enviado` se omite aunque no tenga fecha de comprobación en
Enviados. Un mensaje ausente no se puede dar por entregado solo por ese registro.

## Archivos de salida

| Archivo | Uso |
| --- | --- |
| `*.pdf` | Consentimientos con campos editables de firma, fecha y localidad. |
| `registro_consentimientos.json` | Asociación de filas y PDFs; conservar junto al lote. |
| `emails_lista.csv` | Destinatario, nombre, asunto y archivo adjunto para revisión. |
| `emails_para_enviar.txt` | Texto local de correos; no son borradores de Gmail. |
| `historial_borradores.json` | Estados, identificadores y comprobaciones de envío. |
| `informe_envios.csv` | Aceptación SMTP y comprobación en Enviados, con fechas UTC. |
| `errores_correos.txt` | Último diagnóstico de correo fallido que pudo guardarse. |
| `operacion_correo.lock` | Bloqueo temporal contra operaciones simultáneas. |

## Plantilla y configuración

La plantilla principal es `Docs/plantilla_domiciliacion_sepa.docx`; puedes elegir
otra. El mapeo permite usar encabezados distintos a los marcadores. Campos
habituales: `{{referencia_orden}}`, `{{nombre_deudor}}`, `{{direccion_deudor}}`,
`{{codigo_postal}}`, `{{poblacion}}`, `{{provincia}}`, `{{pais_deudor}}`, `{{iban}}`,
`{{swift_bic}}`, `{{fecha}}` y `{{localidad_firma}}`.
Las fechas de Excel se presentan como `DD/MM/AAAA`; las horas, como `HH:MM`.

Introduce la cuenta y su contraseña de aplicación de Google y pulsa **Guardar
Configuración**. Gestiona la contraseña en [Google](https://myaccount.google.com/apppasswords).
La configuración está en `.sepas_config.json`, dentro del perfil de Windows.

La contraseña se cifra con Fernet usando una clave derivada del nombre de usuario
de Windows. No sustituye un almacén de credenciales del sistema: conocer el nombre
permite reproducir la clave. Existe también compatibilidad con valores sin cifrar
si falla el cifrado. No publiques ese archivo. La contraseña se utiliza para
autenticarse con Gmail mediante TLS y no se incluye en los informes.

## Cambios de septiembre de 2026

- Asociación verificable entre cada fila y su PDF; nombres únicos y limpieza del lote.
- Generación en segundo plano con modal de progreso.
- Envío directo con confirmación, respuesta SMTP y comprobación en Enviados.
- Compatibilidad con el historial anterior para no repetir registros enviados.
- Reconocimiento de `correu` y errores con filas y columnas.
- Apertura del PDF con clic o Intro desde la confirmación.
- Documentación de pruebas y actualización de la entrega en `E:/Sepas/dist`.

## Autor y términos

Autor: [JaviFRx](https://github.com/JaviFRx).

Software propiedad de Javier Fernández Ramos. Uso y distribución sujetos a
autorización expresa del autor. No reproducir, distribuir o modificar sin esa
autorización. Las contraseñas almacenadas son responsabilidad del usuario.
