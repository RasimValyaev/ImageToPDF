# pip install fpdf

# создает pdf и добавляет в него изображения из папки
# в работе использует скрипт ImageCompressed, изменяющий размер изображения
import glob
import os
import re
from fpdf import FPDF
from PIL import Image
from ImageCompressed import compress_img
from pathlib import Path

Image.MAX_IMAGE_PIXELS = None


def add_image_to_pdf(image_directory, save_to_path):
    # path = Path(image_directory)
    # path_name = str(path.absolute())
    # file_name = path.name
    # base_name = path.stem
    extensions = ('*.jpg', '*.jpeg', '*.png', '*.gif')
    pdf = FPDF()
    imagelist = []
    for ext in extensions:
        imagelist.extend(glob.glob(os.path.join(image_directory, ext)))

    for image_file in imagelist:
        cover = Image.open(image_file)
        width, height = cover.size

        # convert pixel in mm with 1px=0.264583 mm
        width, height = float(width * 0.264583), float(height * 0.264583)
        # width = 1240 # 790
        # height = 1754 # 1122
        # given we are working with A4 format size
        pdf_size = {'P': {'w': 210, 'h': 297}, 'L': {'w': 297, 'h': 210}}

        # get page orientation from image size
        orientation = 'P' if width < height else 'L'

        #  make sure image size is not greater than the pdf format size
        width = width if width < pdf_size[orientation]['w'] else pdf_size[orientation]['w']
        height = height if height < pdf_size[orientation]['h'] else pdf_size[orientation]['h']

        pdf.add_page(orientation=orientation)
        image_file = compress_img(image_file)
        pdf.image(image_file, 0, 0, width, height)
        print('size ok')
        pass

    pdf_file = re.split(r'\\', image_directory)[-2] + ' ' + re.split(r'\\', image_directory)[-1]
    save_as = os.path.join(save_to_path, pdf_file + ".pdf")
    pdf.output(save_as, "F")


if __name__ == '__main__':
    image_directory = r"\\PRESTIGEPRODUCT\Scan\ЕСП - Copy\Resize"
    add_image_to_pdf(image_directory, image_directory)
