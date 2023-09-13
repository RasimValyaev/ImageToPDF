from PIL import Image, ImageDraw, ImageFont

# Создаем новое изображение (белый фон)
width, height = 200, 100
background_color = (255, 255, 255)
image = Image.new("RGB", (width, height), background_color)

# Загружаем шрифт
font = ImageFont.truetype("arial.ttf", 24)  # Укажите путь к шрифту

# Создаем объект ImageDraw
draw = ImageDraw.Draw(image)

# Устанавливаем цвет текста и позицию
text_color = (0, 0, 0)  # Черный цвет
text_position = (20, 40)

# Добавляем текст на изображение
text = "СХВАЛЕНО"
draw.text(text_position, text, fill=text_color, font=font)

# Сохраняем изображение
image.save("штамп.png")

# Открываем изображение в стандартном просмотрщике
image.show()
