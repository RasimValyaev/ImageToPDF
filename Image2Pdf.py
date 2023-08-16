# -*- coding: utf-8 -*-

# pip install Pillow
import os

from PIL import Image




def cycle_on_directory_files_and_image_2_pdf(image_path_source):
    # цикл по папкам и файлам в папке pathName
    for root, dirs, files in os.walk(image_path_source):
        try:
            for i, name in enumerate(files):
                filename, file_extension = os.path.splitext(name.lower())
                if file_extension in ['.jpg', '.png', '.bmp'] and filename[:2] in ['рн', 'вн', 'тт']:
                    print(name)
                    image_2_pdf(image_path_source, image_path_source + r'\\' + name)

        except Exception as e:
            print(str(e))
            continue


def image_2_pdf(file_path, file_name):
    # https://auth0.com/blog/image-processing-in-python-with-pillow/
    # изменяет размер изображения
    try:
        # Имя файла
        base_name = os.path.basename(file_name)
        image = Image.open(file_name)
        # new_image = image.resize((1000, 1400))
        new_path = os.path.join(file_path, 'PDF')
        if not os.path.exists(new_path):
            # папки нет. Создаем
            os.makedirs(new_path)

        new_file_path = os.path.join(file_path, 'PDF', base_name)

        # Сохраняем файл в новой папке
        # new_image.save(new_path)
        # image_1 = Image.open(new_path)
        im_1 = image.convert('RGB')
        im_1.save(new_file_path.replace("jpg", "pdf"))

    except Exception as e:
        print(str(e))


if __name__ == '__main__':
    image_path_source = r'C:\Rasim\Scan\Rido\2023-08-15_162928'
    cycle_on_directory_files_and_image_2_pdf(image_path_source)
