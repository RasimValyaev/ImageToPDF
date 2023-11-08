import tabula
import camelot
import pandas as pd

# Provide the path to the PDF file
pdf_file = r"C:\Users\Rasim\Desktop\Fatih bey\FATI_H BEYE O_ZEL FI_YAT C_ALIS_MASI MR.FOOD'S 26.10.2023.pdf"
excel_file = pdf_file.replace(".pdf", ".xlsx")
tables = camelot.read_pdf(pdf_file, flavor="lattice")
# Specify the page number or range to extract (optional)
# pages = 1  # Extract data from page 1
# pages = "1,2,3"  # Extract data from pages 1, 2, and 3
pages = 1
df = tabula.read_pdf(pdf_file)
# Provide the output Excel file name
tabula.convert_into(pdf_file, excel_file, output_format="xlsx", pages=pages)

print(f"PDF data has been converted to {excel_file}")
