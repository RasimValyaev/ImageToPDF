# -*- coding: utf-8 -*-

# программа вытаскивает текст из фото

# pip install Pillow==9.5.0
# pip install torch torchvision torchaudio
# pip install easyocr
# pip install --force-reinstall -v "Pillow==9.5.0"
# pip install pathvalidate
# python -m pip install cyrtranslit # https://github.com/opendatakosovo/cyrillic-transliteration
# pip install psycopg2
import os
import sys
import easyocr
import re
import time
import traceback
from datetime import datetime
import cyrtranslit  # translating ua word to latin. need to the path
import shutil
import psycopg2
from pathvalidate import sanitize_filepath  # deleting incorrect characters in file path
from Image2Pdf import cycle_on_directory_files_and_image_2_pdf, image_2_pdf

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"
sys.path.append(os.path.abspath(CONFIG_PATH))
from configPrestige import username, psw, hostname, port, basename, URL_CONST, chatid_rasim, DATA_AUTH, schema

MONTH = ['СІЧНЯ', 'ЛЮТОГО', 'БЕРЕЗНЯ', 'КВІТНЯ', 'ТРАВНЯ', 'ЧЕРВНЯ', 'ЛИПНЯ', 'СЕРПНЯ', 'ВЕРЕСНЯ', 'ЖОВТНЯ',
         'ЛИСТОПАДА', 'ГРУДНЯ']


def ger_correct_month(mont_name):
    if mont_name == 'ПИПНЯ':
        return 'ЛИПНЯ'
    else:
        return mont_name


def con_postgres_psycopg2():
    conpg = ''

    try:
        conpg = psycopg2.connect(dbname=basename, user=username,
                                 password=psw, host=hostname, port=port)
        conpg.set_client_encoding('UNICODE')

    except Exception as e:
        sms = "ERROR:con_postgres_psycopg2: %s" % e
        print(sms)
        return ''

    finally:
        return conpg


def get_unix_time(timestamp):
    unix_time = time.time()
    cur_time = time.strftime("%d.%m.%Y %H:%M:%S", time.localtime(unix_time))
    cur_time_unix_format = get_unix_time(cur_time)
    print(cur_time_unix_format)
    return int(time.mktime(time.strptime(timestamp, '%d.%m.%Y %H:%M:%S')))


def create_file_name(doc_type, doc_date, doc_number):
    postfix = ''
    if doc_type == 'ТТН2':
        doc_type = 'ТТН'
        postfix = '_1'

    return '{} {} {}{}'.format(doc_type, doc_number, doc_date, postfix)


def get_doc_type(txt_source):
    doc_type = ''
    for item in txt_source:
        item_upper = item.upper()
        if 'ВИДАТКОВА' in item_upper:
            doc_type = 'ВН'
        elif 'ТОВАРНО-ТРАНСПОРТНА НАКЛАДНА' in item_upper:
            doc_type = 'ТТН'
        elif 'ВІДОМОСТІ ПРО ВАНТАЖ' in item_upper:
            doc_type = 'ТТН2'

        if doc_type != '':
            return doc_type


def get_doc_number(txt_source):
    doc_number = 0
    for item in txt_source:
        if 'ng' in item.lower():
            doc_number = re.split(r"\s", item.lower())
            if len(doc_number) == 2:
                doc_number = int(re.search(r"\d+", doc_number[1])[0])
                break

    if doc_number == 0:
        print('Не определился номер док', txt_source)

    return f"{doc_number:05d}"


def get_doc_date(txt_source):
    source = ''
    date_doc = ''
    try:
        for item in txt_source:
            if 'року' in item:
                source = re.split(r"\s", item.lower())
                if len(source) == 4:
                    date = source[-4]
                    month = source[-3].upper()
                    month = ger_correct_month(month)
                    month_index = MONTH.index(month) + 1
                    if date == '98':
                        date = '08'
                    elif date not in range(1, 31):
                        print('Дата должна быть в интервале от 1 до 31', date)
                        break

                    year = source[-2]
                    if int(year) < 2020:
                        print('Год должен быть больше 2020', year)
                        break

                    date_doc = date + '.' + f"{month_index:02d}" + '.' + year
                    date_doc = datetime.strptime(date_doc, '%d.%m.%Y')
                    date_doc = datetime.strftime(date_doc, '%d.%m.%Y')
                    break

        if date_doc == '':
            print('Не нашел данных по дате', txt_source)

    except Exception as e:
        date_doc = ''
        print('ошибка при преобразовании в дату', source)
        print(str(e))

    finally:
        return date_doc


def convert_image2txt(file):
    reader = easyocr.Reader(['uk', 'en'], gpu=False, verbose=False)
    return reader.readtext(file, detail=0)


def get_counterparty(doc_date, doc_number):
    conpg = con_postgres_psycopg2()
    cur = conpg.cursor()
    sql = f'''
        SELECT DISTINCT customer
        FROM public.v_one_sale x
        WHERE (doc_date = '{doc_date}') AND (doc_number LIKE '%{doc_number}%')    
    '''
    cur.execute(sql)
    return cur.fetchone()[0]


def image_read(temp_file_name):
    try:
        result = convert_image2txt(temp_file_name)
        doc_type = get_doc_type(result)
        doc_number = get_doc_number(result)
        doc_date = get_doc_date(result)
        if doc_date == '':
            return ''
        counterparty = get_counterparty(doc_date, doc_number)
        print(counterparty)
        return doc_type, doc_date, doc_number, counterparty

    except Exception as e:
        print(traceback.format_exc())


def image_lists(folder):
    for file_name in os.listdir(folder):
        try:
            count_double = 0  # count double file name
            filename, file_extension = os.path.splitext(file_name.lower())
            if file_extension in ['.jpg', '.png', '.bmp']:
                unix_time = int(time.time())
                temp_file_name = str(unix_time) + file_extension
                print(folder, file_name, temp_file_name)

                # copy image to project path before read
                # path with the cyrillic can't read
                if os.path.isfile(temp_file_name):
                    os.remove(temp_file_name)
                else:
                    old_file = os.path.join(folder, file_name)
                    shutil.copy2(old_file, temp_file_name)

                doc_type, doc_date, doc_number, counterparty = image_read(temp_file_name)
                file_without_path = create_file_name(doc_type, doc_date, doc_number).upper()
                file_without_path = file_without_path + file_extension
                # moving image to counterparty_folder, before converting to pdf
                new_folder = sanitize_filepath(os.path.join(folder, counterparty))
                if ((sanitize_filepath(counterparty) not in os.path.basename(os.path.dirname(folder)))
                        and not os.path.isdir(new_folder)):
                    os.makedirs(sanitize_filepath(new_folder))

                file_with_path = os.path.join(new_folder, file_without_path)

                if os.path.isfile(file_with_path):
                    # os.remove(file_with_path)
                    file_without_path = create_file_name(doc_type, doc_date,
                                                         doc_number) + '_' + str(count_double) + file_extension
                    file_with_path = os.path.join(new_folder, file_without_path)
                    count_double += 1

                os.rename(temp_file_name, file_with_path)
                # convert and save image file to pdf

                image_2_pdf(new_folder, file_with_path)
                os.remove(os.path.join(folder, file_name))

        except Exception as e:
            print(traceback.format_exc())
            print(str(e))
            continue


if __name__ == '__main__':
    image_lists(r'C:\Rasim\Scan\Rido\2023-08-15_162928')
