import os

import PyPDF2


def scan_folder(folder_path):
    for file_name in os.listdir(folder_path):
        try:
            pdf = os.path.join(folder_path, file_name)
            pdf = PyPDF2.PdfFileReader(pdf)
            page0 = pdf.getPage(0)
            page0.scaleBy(1.8)  # float representing scale factor - this happens in-place
            writer = PyPDF2.PdfFileWriter()  # create a writer to save the updated results
            writer.addPage(page0)
            new_file = os.path.join(r"C:\Users\Rasim\Desktop\Scan\ESP_resize\7", file_name)
            with open(new_file, "wb+") as f:
                writer.write(f)

        except Exception as e:
            print(str(e))


if __name__ == '__main__':
    # path_to_folder = r'\\PRESTIGEPRODUCT\Scan\ЕСП'
    path_to_folder = r'C:\Users\Rasim\Desktop\Scan\ESP_resize'
    scan_folder(path_to_folder)
