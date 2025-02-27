from docx import Document

# Se carga el documento de Word
doc = Document('FORMATO DE LISTA DE ASISTENCIA enero 2025.docx')

# Se recorren todas las tablas y se muestra su índice y su contenido
for index, table in enumerate(doc.tables):
    print(f"\nIndice de la tabla: {index}")

    # Se muestra el contenido de las tablas {Solo las primeras filas}
    for row in table.rows[:3]: # Solo se muestran las 3 primeras filas
        print([cell.text.strip() for cell in row.cells])