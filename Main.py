# назначение: в папке отбирает pdf файлы
# создает папки с наименованием = тип документа + дата
# из pdf в эту папку извлекает изображения
# создает 1 pdf файл и все изображения заносит туда
#

import os
import re
import pandas as pd

from datetime import datetime
from pathlib import Path
from Image2PdfMultiPages import add_image_to_pdf
from PdfExtractImage import extract_image
from MergeExcleWord2 import save_df_to_excel, get_pdf_set_with_date_in_file_name, convert_date_to_str_df


# pd.set_option('precision', 2)


# создаем папки по циклу согласно типу_док и дате, и извлекаем туда изображения из pdf
def cycle_for_dates(df_excel, df_pdf):
    try:
        # df_excel = df_excel.sort_values("датаРеализации") # column format hate to datetime
        # df_excel['датаРеализации'] = pd.to_datetime(df_excel['датаРеализации'], format='%d.%m.%Y').dt.strftime('%d.%m.%Y')
        df_pdf = convert_date_to_str_df(df_pdf,'датаРеализации')
        doc_types = df_pdf['doc_type'].unique()
        dates = df_pdf['датаРеализации'].unique()
        doc_number_list = []
        for date in dates:
            for doc_type in doc_types:
                df_result = df_pdf[(df_pdf['датаРеализации'] == date) & (df_pdf['doc_type'] == doc_type)]
                image_save_to_path = ''
                for i, row in df_result.iterrows():

                    date_revers = datetime.strptime(date, "%d.%m.%Y").strftime("%Y.%m.%d")
                    image_save_to_path = os.path.join(os.path.dirname(row.filename), date_revers, doc_type)
                    print(row.filename)
                    if not os.path.exists(image_save_to_path):  # the folder create here, because we're using row
                        os.makedirs(image_save_to_path)

                    extract_image(row.filename, image_save_to_path)  # extract images from pdf to image_save_to_path
                doc_number_list.append(df_result['номерРеализации'].values.tolist())
                if os.path.exists(image_save_to_path):
                    # pdf_save_to_path = Path(image_save_to_path,"Group").parents[1]
                    pdf_save_to_path = Path(image_save_to_path, "Group")
                    add_image_to_pdf(image_save_to_path, pdf_save_to_path)  # add image to pdf
                    # os.remove(Path(image_save_to_path).parents[0])

                # merge_excel_and_word(excel_file_source)

                print('**************************\n', doc_type, date)

    except Exception as e:
        err_info = "Error: Main: %s" % e
        print(err_info)


if __name__ == '__main__':
    extension = ['*.pdf']
    excel_file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    df_excel, df_pdf = save_df_to_excel(excel_file_source)
    # pdf_directory = os.path.dirname(excel_file_source)
    # df_pdf = get_pdf_set_with_date_in_file_name(pdf_directory)
    if len(df_excel) > 0 and len(df_pdf) > 0:
        cycle_for_dates(df_excel, df_pdf)
