# Guía rápida

## Generar, revisar y enviar

1. Abre `E:/Sepas/dist/Generador_Consentimientos_SEPAS.exe`.
2. Selecciona el Excel `.xlsx`, la plantilla Word y la carpeta de salida.
3. Revisa el mapeo y genera los PDFs. Espera a que termine la modal.
4. Introduce la cuenta Gmail y su contraseña de aplicación; guarda la configuración.
5. Pulsa **Configurar correo** y revisa asunto, texto y `emails_lista.csv`.
6. Pulsa **Enviar correos (Gmail)**.
7. **Haz clic en una fila para abrir su PDF**. También puedes seleccionarla y pulsar Intro.
8. Confirma **Enviar los N correos** cuando termines de revisar.
9. Consulta el resumen y `informe_envios.csv` para ver el detalle.

Abrir un PDF no envía el correo. Cancelar la confirmación tampoco envía nada.
Si el PDF cambió o ya no existe, se muestra un error al intentar abrirlo.

## Antes de generar otro lote

La generación borra los PDFs anteriores y las listas locales de correo de la
carpeta seleccionada. Conserva historiales e informes. Regenerar crea un lote
nuevo: revisa qué se envió antes. No borres `registro_consentimientos.json`
ni `historial_borradores.json`.

## Si aparece un error

| Aviso o situación | Qué hacer |
| --- | --- |
| Falta el destinatario | Revisa la fila y columnas indicadas; el correo está vacío. |
| Varios correos distintos | Deja una sola dirección de destino para esa fila. |
| Correo inválido o varias direcciones | Escribe una dirección válida por celda de correo. |
| Columna no reconocida | Usa un encabezado con `email`, `mail`, `correo` o `correu`. |
| Excel o PDF cambiado | Revisa los datos y genera un lote completo con ese Excel. |
| PDF bloqueado al generar | Cierra el PDF y vuelve a generar. |
| Aceptado, pero no comprobado en Enviados | El lote se detuvo. Revisa Gmail; no reenvíes ese correo a ciegas. |
| Envío sin confirmar | Puede haberse enviado antes de perder la conexión. Comprueba Gmail antes de resolver el estado. |
| Otra operación de correo abierta | Espera. Si hubo un cierre inesperado, revisa Gmail e historial antes de retirar `operacion_correo.lock`. |
| Credenciales inválidas | Usa una contraseña de aplicación de Google y comprueba la cuenta. |

Los errores muestran hoja y fila real del Excel, incluso con filas vacías
intermedias. El detalle se guarda en `errores_correos.txt` cuando es posible.
Si corriges datos del Excel, vuelve a generar los PDFs antes de enviar.
Se admite `Direcció correu electrònic (minúscula)` sin renombrar la columna.

## Interpretar el resumen

- **Aceptados por Gmail:** Gmail aceptó el correo para enviarlo.
- **Comprobados en Enviados:** se localizó una copia con su destinatario y PDF.
- **Ya enviados anteriormente:** el historial evita repetirlos.
- **Pendientes de enviar:** quedan mensajes sin aceptación registrada en esta
  operación. Si hubo un fallo de conexión, lee el aviso antes de reintentarlo.

Estos estados no confirman que el destinatario lo haya recibido o leído.

## Borradores antiguos

La versión actual envía directamente y no modifica Borradores en Gmail. No envíes
manualmente un borrador antiguo si la aplicación ya envió ese correo.

Consulta [INSTRUCCIONES_EXE.md](INSTRUCCIONES_EXE.md) para actualizaciones y
[README.md](README.md) para el funcionamiento completo.

Software propiedad de Javier Fernández Ramos. Uso sujeto a autorización del autor.
