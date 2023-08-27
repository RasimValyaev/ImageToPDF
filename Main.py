# назначение: в папке отбирает pdf файлы
# создает папки с типом документа и датой
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


# pdf set with date in file name
def get_pdf_set_with_date_in_file_name(directory):
    ext = r'\d{2}.\d{2}.\d{4}.pdf$'
    data = {}
    doc_type_list = []
    date_list = []
    file_list = []
    for filename in os.listdir(directory):
        if re.search(ext, filename):
            date = os.path.splitext(re.search(ext, filename).group(0))[0]
            date_list.append(date)
            file_list.append(os.path.join(directory, filename))
            doc_type_list.append(re.search('[а-яА-ЯёЁa-zA-Z]+', filename)[0])
            data.update({"doc_type": doc_type_list, "date": date_list, "filename": file_list})

    df = pd.DataFrame(data)
    return df


# создаем папки по циклу согласно типу_док и дате, и извлекаем туда изображения из pdf
def cycle_for_dates(df):
    df['date'] = pd.to_datetime(df['date'],
                                format='%d.%m.%Y')  # if we don't use the "format", program displays a message
    df = df.sort_values("date")
    df['date'] = pd.to_datetime(df['date'], format='%d.%m.%Y').dt.strftime('%d.%m.%Y')
    doc_types = df['doc_type'].unique()
    dates = df['date'].unique()
    for date in dates:
        for doc_type in doc_types:
            df_result = df[(df['date'] == date) & (df['doc_type'] == doc_type)]
            image_save_to_path = ''
            for i, row in df_result.iterrows():
                date_revers = datetime.strptime(date, "%d.%m.%Y").strftime("%Y.%m.%d")
                image_save_to_path = os.path.join(os.path.dirname(row.filename), date_revers, doc_type)
                print(row.filename)
                if not os.path.exists(image_save_to_path):  # the folder create here, because we're using row
                    os.makedirs(image_save_to_path)

                extract_image(row.filename, image_save_to_path)

            if os.path.exists(image_save_to_path):
                pdf_save_to_path = Path(image_save_to_path).parents[1]
                add_image_to_pdf(image_save_to_path, pdf_save_to_path)  # add image to pdf
                # os.remove(Path(image_save_to_path).parents[0])

            print('**************************\n', doc_type, date)


if __name__ == '__main__':
    extension = ['*.pdf']
    pdf_directory = r"\\PRESTIGEPRODUCT\Scan\ЕСП - Copy"
    df = get_pdf_set_with_date_in_file_name(pdf_directory)  # df with date, filename
    cycle_for_dates(df)
