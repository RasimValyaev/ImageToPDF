# Initialize
import os
from PyPDF2 import PdfFileWriter, PdfFileReader


def folder_exists(full_path):
    if not os.path.exists(full_path):
        os.makedirs(full_path)


def scan_folder(folder_path):
    for file_name in os.listdir(folder_path):
        try:
            current_pdf = os.path.join(folder_path, file_name)
            input1 = PdfFileReader(open(current_pdf, "rb"))

            # Make a simple list of page objects
            pages = []
            for i in range(input1.getNumPages()):
                pages.append(input1.getPage(i))

            output = PdfFileWriter()
            # Scale and add the pages to the output object
            count = 1
            for page in pages:
                page.scaleTo(width=8.5, height=11.0)
                output.addPage(page)
                print("Page %d is done!" % (count))
                count += 1

            copy_to_folder = os.path.join(folder_path, "Resize")
            folder_exists(copy_to_folder)
            new_file = os.path.join(copy_to_folder, file_name)

            # Make and write to an output document
            out_doc = open(new_file, 'wb')
            output.write(out_doc)
            out_doc.close()

        except Exception as e:
            print(str(e))


if __name__ == '__main__':
    path_to_folder = r'\\PRESTIGEPRODUCT\Scan\ЕСП - Copy'
    scan_folder(path_to_folder)
