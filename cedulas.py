import openpyxl

# 1. Cargar el documento
archivo = "LEGO OPERACIONES AGO 26.xlsx"
wb = openpyxl.load_workbook(archivo)
ws_lista = wb["Lista Elementos"]
ws_info = wb["Elementos Info"]
ws_destino = wb["LEGO OPERACIONES"]

# 2. Extraer la lista de nombres requeridos
nombres_requeridos = []
# Asumiendo que los nombres están en la columna A empezando en la fila 2
for row in range(2, ws_lista.max_row + 1):
    nombre = ws_lista.cell(row=row, column=1).value
    if nombre:
        nombres_requeridos.append(str(nombre).strip().split()) # Separar por palabras

# 3. Mapear los bloques en "Elementos Info" por persona
bloques_info = []
bloque_actual = []
nombre_actual = None

for row in range(1, ws_info.max_row + 1):
    celda_nombre = ws_info.cell(row=row, column=6) # Columna F
    
    # Si la celda tiene valor y la fuente está en negrita, es una nueva persona
    if celda_nombre.value and celda_nombre.font and celda_nombre.font.bold:
        if bloque_actual:
            bloques_info.append({"nombre": nombre_actual, "filas": bloque_actual})
        nombre_actual = str(celda_nombre.value).strip().split()
        bloque_actual = [row]
    elif nombre_actual:
        # Sigue agregando filas al bloque de la persona actual
        bloque_actual.append(row)

# Guardar el último bloque leído
if bloque_actual:
    bloques_info.append({"nombre": nombre_actual, "filas": bloque_actual})

# 4. Filtrar y pegar en "LEGO OPERACIONES"
fila_destino = 22

for persona in bloques_info:
    # Comprobar si todas las palabras del nombre requerido están en el nombre de Elementos Info
    es_requerido = any(all(palabra in persona["nombre"] for palabra in req) for req in nombres_requeridos)
    
    if es_requerido:
        # Copiar cada fila del bloque
        for fila_origen in persona["filas"]:
            for col in range(1, ws_info.max_column + 1):
                valor_origen = ws_info.cell(row=fila_origen, column=col).value
                ws_destino.cell(row=fila_destino, column=col).value = valor_origen
            fila_destino += 1
            
        # Añadir la fila en blanco de separación al terminar el bloque de la persona
        fila_destino += 1

# Guardar el resultado en un nuevo archivo
wb.save("LEGO_OPERACIONES_ACTUALIZADO.xlsx")
print("Extracción y copiado completado con éxito.")