# pip install python-docx Pillow

# экспортирует каждую строчку в изображение
import docx
from PIL import Image, ImageDraw, ImageFont


def convert_to_png(filename):
    doc = docx.Document(filename)
    for i, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text
        if text:
            image = Image.new('RGB', (1000, 500), (255, 255, 255))
            draw = ImageDraw.Draw(image)
            font = ImageFont.truetype('arial.ttf', 16)
            draw.text((10, 10), text, font=font, fill=(0, 0, 0))
            image.save(f'output_{i}.png')


if __name__ == '__main__':
    file_name = r'C:\Rasim\Python\Prestige\TelegramBot\001694007370.docx'
    convert_to_png(file_name)
