import openpyxl
from fpdf import FPDF

# Define variable to load the dataframe
excel_dataframe = openpyxl.load_workbook("autoventa.xlsx")

# Define variable to read sheet
dataframe = excel_dataframe.active

data = []

# Iterate the loop to read the cell values
for row in range(1, dataframe.max_row):
    _row = [row,]

    for col in dataframe.iter_cols(1, dataframe.max_column):
        _row.append(col[row].value)

    data.append(_row)

headers = ["#","Codigo", "Descripcion", "Unidades", "Total"]

# Create a new PDF in landscape orientation
pdf = FPDF(orientation='L', unit='mm', format='A4')
pdf.add_page()
pdf.set_font("Arial", size=8)

# Calculate the width of each cell to fit the page width
page_width = pdf.w - 2 * pdf.l_margin  # Page width minus margins
col_width = page_width / len(headers)  # Equal width for each column

# Add the header row
for header in headers:
    pdf.cell(col_width, 10, header, 1, 0, 'C')
pdf.ln()

# Add the data rows
for row_data in data:
    for cell_value in row_data:
        text = str(cell_value) if cell_value is not None else ''
        # Ensure the text is encoded as UTF-8
        pdf.cell(col_width, 10, text.encode('latin-1', 'replace').decode('latin-1'), 1, 0, 'C')
    pdf.ln()

# Save the PDF using UTF-8 encoding
pdf.output("output.pdf", 'F')