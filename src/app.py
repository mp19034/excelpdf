import openpyxl
from fpdf import FPDF

# Cargar el archivo Excel
excel_dataframe = openpyxl.load_workbook("autoventa.xlsx")

# Seleccionar la hoja activa del archivo Excel
dataframe = excel_dataframe.active

data = []

# Iterar para leer los valores de las celdas
for row in range(1, dataframe.max_row):
    _row = [row,]
    
    # Iterar sobre las columnas y agregar los valores al row
    for col in dataframe.iter_cols(1, dataframe.max_column):
        _row.append(col[row].value)
    
    # Validar y calcular según las condiciones especificadas
    #
    #
    #
    
    if len(_row) > 2 and _row[1] == 11001:
        if _row[3] < 54:
            _row.append(_row[3])  # Agregar a la sexta columna
            _row.insert(5, 0)  # Espacio para la quinta columna
        else:
            _row.append(0)  # Espacio para la sexta columna
            _row.insert(5, _row[3]/54)  # Agregar a la quinta columna
    else:
        _row.append(None)  # Espacio para la sexta columna
        _row.insert(5, 0)  # Espacio para la quinta columna
    
    data.append(_row)
    
    #
    #
    #
    #
    
# Encabezados de las columnas en el PDF
headers = ["#", "Codigo", "Descripcion", "Unidades", "UMB", "Cajas Completas", "Unidades"]

# Texto que deseas agregar encima de la tabla
texto_superior = "Productos Alimenticios Diana, S.A de C.V. \n Distribuidora Santa Ana \n Picking List"

# Texto inferior con tabuladores para alinear Ruta: a la izquierda y Picking: a la derecha
ruta_texto = "Ruta:"
picking_texto = "Picking:"
line = "_________________________________________________________________________________________________________________________________________________________________________________________________________________________________________"
vendedor = "Vendedor:"
carga = "Carga de Mercancia:"
pickig_list = "Picking List:"

# Crear un nuevo PDF en orientación horizontal (landscape)
pdf = FPDF(orientation='L', unit='mm', format='A4')
pdf.add_page()
pdf.set_font("Arial", size=6)

# Calcular el ancho de cada celda para que ajuste al ancho de la página
page_width = pdf.w - 2 * pdf.l_margin  # Ancho de la página menos los márgenes

# Agregar texto encima de los encabezados y la tabla
pdf.set_xy(20, 5)  # Posición XY para el texto superior
pdf.multi_cell(page_width - 20, 8, texto_superior, 0, 'C')
pdf.multi_cell(page_width , 2, line, 0, 'C')

# Calcular la altura del texto superior
texto_superior_height = pdf.get_y()

# Agregar texto inferior con tabuladores para alinear Ruta: a la izquierda y Picking: a la derecha
pdf.set_xy(10, texto_superior_height + 5)  # Posición XY para el texto inferior
pdf.cell(page_width // 2 - 20, 10, ruta_texto, 0, 0, 'L')
pdf.cell(page_width // 2 - 10, 10, picking_texto, 0, 0, 'R')

#Vendedor y carga
pdf.set_xy(10, texto_superior_height + 15)  # Posición XY para el texto inferior
pdf.cell(page_width // 2 - 20, 10, vendedor, 0, 0, 'L')
pdf.cell(page_width // 2 - 10, 10, carga, 0, 0, 'R')

#pickin list
pdf.set_xy(10, texto_superior_height + 25)  # Posición XY para el texto inferior
pdf.cell(page_width // 2 - 20, 10, pickig_list, 0, 0, 'L')

# Agregar la fila de encabezados
pdf.set_xy(pdf.l_margin, pdf.get_y() + 10)  # Posicionar debajo del texto superior
for header in headers:
    pdf.cell(page_width / len(headers), 6, header, 1, 0, 'C')
pdf.ln()

# Agregar las filas de datos
for row_data in data:
    for cell_value in row_data:
        text = str(cell_value) if cell_value is not None else ''
        # Asegurar que el texto esté codificado como UTF-8
        pdf.cell(page_width / len(headers), 6, text.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
    pdf.ln()

# Guardar el PDF utilizando codificación UTF-8
pdf.output("output.pdf", 'F')
