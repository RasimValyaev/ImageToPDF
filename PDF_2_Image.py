# pip install pdf2image
# download folder from https://github.com/oschwartz10612/poppler-windows/releases/ and paste to current dir
import os
from pdf2image import convert_from_path
from Image2Txt import image_read


def read_pdf(filename):
    path_to_bin = r'c:\Rasim\Python\ImageToPDF\poppler-23.08.0\Library\bin'
    images = convert_from_path(filename, poppler_path=path_to_bin, dpi=200)
    x = 1
    for image in images:
        try:
            new_filename = r'C:\Users\Rasim\Desktop\Scan\7\deneme_' + str(x) + '.jpeg'
            image.save(new_filename, 'JPEG')
            doc_type, doc_date, doc_number, counterparty = image_read(new_filename)
            print(doc_type, doc_date, doc_number, counterparty)

        except Exception as e:
            print(str(e))

        finally:
            break


def scan_folder(folder_path):
    for file_name in os.listdir(folder_path):
        try:
            file_with_path = os.path.join(folder_path, file_name)
            print(file_with_path)
            read_pdf(file_with_path)
        except Exception as e:
            print(str(e))


if __name__ == '__main__':
    path_to_folder = r'C:\Users\Rasim\Desktop\Scan\Amik\PDF'
    scan_folder(path_to_folder)
