# pip install 'PyPDF2<3.0'
# pdf with images does not work
import os
import PyPDF2


def create_folder_if_not_exists(full_path):
    if not os.path.exists(full_path):
        os.makedirs(full_path)


def scan_folder(folder_path):
    new_file = ''
    copy_to_folder = os.path.join(folder_path, "Resize")
    create_folder_if_not_exists(copy_to_folder)
    writer = PyPDF2.PdfFileWriter()  # create a writer to save the updated results
    for file_name in os.listdir(folder_path):
        try:
            pdf = os.path.join(folder_path, file_name)
            pdf = PyPDF2.PdfFileReader(pdf)
            page0 = pdf.getPage(0)
            page0.scaleBy(0.5)  # float representing scale factor - this happens in-place
            new_file = os.path.join(copy_to_folder, file_name)
            writer.addPage(page0)

        except Exception as e:
            print(str(e))

    if new_file != '':
        with open(new_file, "wb+") as f:
            writer.write(f)


if __name__ == '__main__':
    # path_to_folder = r'\\PRESTIGEPRODUCT\Scan\ЕСП'
    path_to_folder = r"C:\Users\Rasim\Desktop\Scan\Маркет позитив плюс"
    scan_folder(path_to_folder)
