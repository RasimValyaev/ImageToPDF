# pip install pytesseract

from PIL import Image, ImageEnhance
import pytesseract
import cv2
image_name = r'C:\Users\Rasim\Desktop\ЕСП\7\2023-08-08_143136\ВН 17386 05.07.2023.jpg'
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
# image = Image.open(image_name)
img=cv2.imread(image_name)
string=pytesseract.image_to_string(img)

# # image.load()
# # Повышенние резкости изображения:
# enhancer1 = ImageEnhance.Sharpness(image)
# factor1 = 0.01  # чем меньше, тем больше резкость
# im_s_1 = enhancer1.enhance(factor1)
# print(pytesseract.image_to_string(image, lang='rus'))
# # text = pytesseract.image_to_string(im_s_1, config='--psm 6 -c tessedit_char_whitelist=0123456789,. ').split(
# #     '\n')
# # print(text)
