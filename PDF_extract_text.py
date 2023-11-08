# importing required modules
from PyPDF2 import PdfReader
from pathlib import Path
pdf_file = r"C:\Users\Rasim\Desktop\Fatih bey\FATI_H BEYE O_ZEL FI_YAT C_ALIS_MASI MR.FOOD'S 26.10.2023.pdf"
# creating a pdf reader object
reader = PdfReader(pdf_file)

# printing number of pages in pdf file
print(len(reader.pages))

# getting a specific page from the pdf file
page = reader.pages[0]

# extracting text from page
text = page.extract_text()
print(text)

file_path = Path(pdf_file.replace(".pdf",".txt"))

# Open the file in write mode and write the text
with open(file_path, "w",  encoding="utf-8") as file:
    file.write(text)
