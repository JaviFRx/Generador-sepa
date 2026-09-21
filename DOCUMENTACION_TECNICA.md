# Documentación técnica

## Componentes

| Archivo | Responsabilidad |
| --- | --- |
| `app_consentimientos.py` | Tkinter, mapeo, modales, revisión de PDFs y confirmación. |
| `modulos/lector_excel.py` | Hoja activa, fechas y seguimiento de filas físicas. |
| `modulos/generador_word.py` | Plantilla Word, conversión a PDF y campos editables. |
| `modulos/registro_consentimientos.py` | Manifiesto y asociación verificable de filas y PDFs. |
| `modulos/preparador_emails.py` | Correos, listas locales y MIME; conserva funciones de borradores para compatibilidad. |
| `modulos/envio_correos.py` | Envío directo, aceptación SMTP, comprobación IMAP e informe. |
| `modulos/envio_borradores.py` | Historial, bloqueo y validación MIME compartidos; contiene además el servicio anterior que la interfaz ya no utiliza. |

## Generación y asociación

El manifiesto contiene versión, indicador `completo`, huellas de registros
originales en orden y entradas por PDF con número, nombre y SHA-256. El mapeo de
la plantilla no sustituye al registro original usado para asociar el adjunto.

Se invalida antes de limpiar la salida y solo se completa si están todos los
documentos. La limpieza elimina PDFs y listas locales, sin recorrer subcarpetas.
Se rechazan rutas fuera de la carpeta. Los nombres incorporan UUID para evitar
sobrescrituras; la conversión comprueba retorno, existencia y tamaño del PDF.
Se valida todo el lote y se cargan sus bytes antes de construir los adjuntos.

## Preparación y confirmación

`EnviadorCorreos.preparar_envio` no conecta con Gmail. Construye un plan con cuenta,
destinatarios, archivos, MIME y huellas. Usa `make_msgid` para que los mensajes
nuevos no reutilicen identificadores de borradores anteriores.

La confirmación asocia cada identificador del Treeview a su entrada del plan.
El clic abre la fila bajo el puntero, no la selección anterior. Antes de
`os.startfile` comprueba carpeta, extensión PDF, existencia y SHA-256. Intro
abre la fila enfocada. Abrir un archivo no inicia el envío.

Tras confirmar se vuelve a validar el lote completo y cada mensaje: remitente,
destinatario, adjunto único y contenido. Se rechazan CC/CCO, cabeceras de reenvío
y alteraciones del plan. Los trabajadores usan una cola que Tk consume con
`after`; no actualizan widgets directamente.

## Historial y envío

Se conserva `historial_borradores.json`, versión 1, con registros por cuenta
normalizada y nombre de PDF. Se escribe mediante archivo temporal y reemplazo.
`operacion_correo.lock` se crea de forma exclusiva durante cada operación; un
cierre inesperado puede dejarlo pendiente de revisión manual.

| Estado | Significado |
| --- | --- |
| `borrador` | Registro preparado o heredado, sin aceptación registrada. |
| `sin_confirmar` | Persistido antes de enviar; una interrupción bloquea repeticiones. |
| `enviado` | Gmail aceptó el correo; se omite para esa cuenta y archivo. |

El servicio conecta por TLS con `imap.gmail.com:993`, identifica la carpeta por
su atributo especial Sent y la selecciona en modo de solo lectura. Después
autentica en `smtp.gmail.com:465`. `SMTPConRegistro` captura la respuesta a DATA,
sin registrar credenciales.

Tras la aceptación se guardan `message_id_envio`, `huella_correo`, `aceptado_utc`,
`smtp_codigo` y `smtp_respuesta`. Si falla la persistencia se detiene el lote y
permanece el último estado durable; no se reintenta automáticamente.

La comprobación busca el Message-ID en Enviados, recupera con `BODY.PEEK[]` y
valida remitente, destinatario, asunto y PDF. Hace hasta tres consultas, con
esperas de 2 y 4 segundos. Si encuentra la copia, guarda `enviados_utc` y
`uid_enviados`. El UID corresponde a esa carpeta en ese momento; el Message-ID
se conserva para búsquedas posteriores.

Si no puede comprobar la copia, mantiene la aceptación, avisa y detiene el lote.
El servicio actual no usa APPEND, STORE ni EXPUNGE ni elimina mensajes de Gmail.
Los fallos al cerrar la conexión no descartan los resultados ya registrados.

El CSV contiene los registros de la cuenta y separa aceptación SMTP y comprobación
en Enviados. Si no puede escribirse, se avisa y el JSON conserva los estados.
La protección no detecta todos los duplicados entre regeneraciones, carpetas
distintas o envíos manuales. Tampoco acredita entrega al destinatario.

## Diagnóstico del Excel

`LectorExcel.filas_excel` conserva números físicos sin añadir claves a los
registros originales ni alterar sus huellas. Los encabezados reconocen `mail`,
`correo` y `correu`. Se distinguen columnas ausentes, valores vacíos, formatos
inválidos y varias direcciones distintas.

Se recogen todos los errores antes de preparar mensajes. Si se puede escribir
`errores_correos.txt`, la ventana muestra hasta ocho y remite al informe completo.
Si no, muestra todos. Estos errores bloquean el envío antes de conectar.

## Pruebas

```powershell
.venv/Scripts/python.exe -m unittest discover -s tests -v
```

Las pruebas usan directorios temporales, cuentas `example.test` y transportes
simulados; no envían correos reales. Las pruebas gráficas necesitan Tcl/Tk:
si no está disponible se omiten, por lo que hay que revisar el resumen.
Las rutas locales están en [INSTRUCCIONES_EXE.md](INSTRUCCIONES_EXE.md).

La batería cubre asociación de PDFs, Excel modificado, generación incompleta,
destinatarios ambiguos, `correu`, filas vacías, rechazos y timeouts de Gmail,
persistencia y reintentos, copias ausentes o incorrectas en Enviados, retardos
de visibilidad y respuesta de las modales. Conserva regresiones del servicio
anterior, incluida la consulta CAPABILITY después del login IMAP.

La apertura por clic se comprobó además con el visor simulado: selección anterior
distinta, PDF cambiado y clic en zona vacía. Las verificaciones locales no
sustituyen una prueba de entrega con destinatarios controlados por el usuario.
