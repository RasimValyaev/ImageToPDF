# pip install 'PyPDF2<3.0'
import os
import PyPDF2


def folder_exists(full_path):
    if not os.path.exists(full_path):
        os.makedirs(full_path)


def scan_folder(folder_path):
    for file_name in os.listdir(folder_path):
        try:
            pdf = os.path.join(folder_path, file_name)
            pdf = PyPDF2.PdfFileReader(pdf)
            page0 = pdf.getPage(0)
            page0.scaleBy(0.5)  # float representing scale factor - this happens in-place
            writer = PyPDF2.PdfFileWriter()  # create a writer to save the updated results
            writer.addPage(page0)
            copy_to_folder = os.path.join(folder_path, "Resize")
            folder_exists(copy_to_folder)
            new_file = os.path.join(copy_to_folder, file_name)
            with open(new_file, "wb+") as f:
                writer.write(f)

        except Exception as e:
            print(str(e))


if __name__ == '__main__':
    # path_to_folder = r'\\PRESTIGEPRODUCT\Scan\ЕСП'
    path_to_folder = r'\\PRESTIGEPRODUCT\Scan\ЕСП - Copy'
    scan_folder(path_to_folder)
