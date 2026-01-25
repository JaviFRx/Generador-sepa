from docx import Document
from docx.shared import Pt, RGBColor

# Abrir la plantilla existente
doc = Document('d:/Sepas/docs/plantilla_domiciliacion_sepa.docx')

# Buscar y reemplazar el campo de referencia
for para in doc.paragraphs:
    if 'Referencia' in para.text and 'orden' in para.text and 'domiciliación' in para.text:
        # Encontramos el párrafo, ahora lo modificamos
        texto_original = para.text
        # Reemplazar la línea de guiones por el marcador
        nuevo_texto = texto_original.replace('______________________________________________________________', '{{nombre_alumno}}_Robotica')
        
        # Limpiar el párrafo
        for run in para.runs:
            run.text = ""
        
        # Agregar el nuevo texto
        if para.runs:
            para.runs[0].text = nuevo_texto
        else:
            para.add_run(nuevo_texto)
        
        print(f'✓ Campo modificado:')
        print(f'  Antes: {texto_original[:80]}...')
        print(f'  Después: {nuevo_texto[:80]}...')
        break

# Guardar la plantilla modificada
doc.save('d:/Sepas/docs/plantilla_domiciliacion_sepa_nueva.docx')
print('\n✓ Plantilla actualizada guardada como: plantilla_domiciliacion_sepa_nueva.docx')
print('  Cierra el archivo original y renómbralo manualmente')
