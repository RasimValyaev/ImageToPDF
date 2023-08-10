# программа переводит фото в текст

# pip install pillow
# pip install pytesseract

# pip install torch torchvision torchaudio
# pip install easyocr
# pip install --force-reinstall -v "Pillow==9.5.0"
import os
import easyocr
import re
import time

MONTH = ['СІЧНЯ', 'ЛЮТОГО', 'БЕРЕЗНЯ', 'КВІТНЯ', 'ТРАВНЯ', 'ЧЕРВНЯ', 'ЛИПНЯ', 'СЕРПНЯ', 'ВЕРЕСНЯ', 'ЖОВТНЯ',
         'ЛИСТОПАДА', 'ГРУДНЯ']


def get_unix_time(timestamp):
    unix_time = time.time()
    cur_time = time.strftime("%d.%m.%Y %H:%M:%S", time.localtime(unix_time))
    cur_time_unix_format = get_unix_time(cur_time)
    print(cur_time_unix_format)
    return int(time.mktime(time.strptime(timestamp, '%d.%m.%Y %H:%M:%S')))


def get_doc_type(txt_source):
    doc_type = ''
    if 'ТОВАРНО-ТРАНСПОРТНА НАКЛАДНА' in txt_source:
        doc_type = 'ТТН'
    elif 'ВІДОМОСТІ ПРО ВАНТАЖ' in txt_source:
        doc_type = 'ТТН2'
    else:
        txt_source = re.split("\s", txt_source[0].upper())
        if txt_source[0] == 'ВИДАТКОВА':
            doc_type = 'ВН'

    return doc_type


def get_doc_number(txt_source, doc_type):
    doc_number = ''
    if doc_type == 'ВН':
        txt_source = re.split("\s", txt_source[0])
        doc_number = txt_source[txt_source.index('Ng') + 1]
    elif doc_type == 'ТТН':
        doc_number = txt_source[5][3:]
    elif doc_type == 'ТТН2':
        txt_source = re.split("\s", txt_source[29])
        doc_number = txt_source[txt_source.index('Ng') + 1]

    return int(re.search("\d+", doc_number)[0])


def get_counterparty(txt_source, doc_type):
    if doc_type == 'ВН':
        return txt_source[8]
    elif doc_type == 'ТТН':
        return txt_source[46]
    elif doc_type == 'ТТН2':
        return re.split("\s", txt_source[111])[0]


def get_doc_date(txt_source, doc_type):
    source = ''
    if doc_type == 'ВН':
        source = re.split("\s", txt_source[0])
    elif doc_type == 'ТТН':
        source = re.split("\s", txt_source[6])
    elif doc_type == 'ТТН2':
        source = re.split("\s", txt_source[32])

    if source != '':
        date = source[-4]
        month = source[-3].upper()
        month_index = MONTH.index(month) + 1
        year = source[-2]
        return date + '.' + f"{month_index:02d}" + '.' + year
    else:
        return ''


def create_file_name(doc_type, get_doc_date, get_doc_number):
    postfix = ''
    if doc_type == 'ТТН2':
        doc_type = 'ТТН'
        postfix = '_1'

    return '{} {} {}{}'.format(doc_type, get_doc_number, get_doc_date, postfix)


def convert_image2txt(file):
    reader = easyocr.Reader(['uk', 'en'])
    return reader.readtext(file, detail=0)


def image_lists(folder):
    for file_name in os.listdir(folder):
        filename, file_extension = os.path.splitext(file_name.lower())
        if file_extension in ['.jpg', '.png', '.bmp']:
            print(folder, file_name)
            unix_time = int(time.time())
            temp_file_name = os.path.join(folder, str(unix_time) + file_extension)
            if os.path.isfile(temp_file_name):
                os.remove(temp_file_name)
            else:
                old_file = os.path.join(folder, file_name)
                os.rename(old_file, temp_file_name)

            result = convert_image2txt(temp_file_name)
            doc_type = get_doc_type(result)
            doc_date = get_doc_date(result, doc_type)
            doc_number = get_doc_number(result, doc_type)
            counterparty = get_counterparty(result, doc_type)
            print(counterparty)
            old_file = os.path.join(folder, temp_file_name)
            new_name_raw = create_file_name(doc_type, doc_date, doc_number) + file_extension
            new_file = os.path.join(folder, new_name_raw)

            if os.path.isfile(new_file):
                os.remove(new_file)
            else:
                os.rename(old_file, new_file)


if __name__ == '__main__':
    image_lists(r'C:\Rasim\Python\ImageToPDF')
