# pip install fpdf
# создает pdf и добавляет в него изображения из папки
# в работе использует другой скрипт, изменяющий размер фото
import glob
import os
from fpdf import FPDF
from PIL import Image
from ImageCompressed import compress_img

Image.MAX_IMAGE_PIXELS = None
image_directory = r"\\PRESTIGEPRODUCT\Scan\ЕСП - Copy\Resize"
extensions = ('*.jpg', '*.jpeg', '*.png', '*.gif')
pdf = FPDF()
imagelist = []
for ext in extensions:
    imagelist.extend(glob.glob(os.path.join(image_directory, ext)))

for imageFile in imagelist:
    cover = Image.open(imageFile)
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
    imageFile = compress_img(imageFile)
    pdf.image(imageFile, 0, 0, width, height)
    print('size ok')
    pass


save_as = os.path.join(image_directory, "file.pdf")
pdf.output(save_as, "F")
