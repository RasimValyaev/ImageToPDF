# назначение: в папке отбирает pdf файлы
# создает папки с наименованием = тип документа + дата
# из pdf в эту папку извлекает изображения
# создает 1 pdf файл и все изображения заносит туда
#

import os
import re
import pandas as pd
from mailmerge import MailMerge
from datetime import datetime
from pathlib import Path
from Image2PdfMultiPages import add_image_to_pdf
from PdfExtractImage import extract_image
from MergeExcleWord import save_df_to_excel, get_pdf_set_with_date_in_file_name, convert_date_to_str_df

NUMBER_FIRST = 1


# создаем папки по циклу согласно типу_док и дате, и извлекаем туда изображения из pdf
def cycle_for_dates(df_exl, df_pdf):
    try:
        df_pdf = convert_date_to_str_df(df_pdf, 'датаРеализации')
        doc_types = df_pdf['doc_type'].unique()
        client_okpo_list = df_exl['Податковий_номер_Покупця'].unique()
        for client_okpo in client_okpo_list:
            df_exl_filtr_okpo = df_exl[(df_exl['Податковий_номер_Покупця'] == client_okpo)].reset_index(drop=True)
            dates = df_exl_filtr_okpo['датаРеализации'].unique().tolist()
            i = 1
            for date in dates:
                record_number = str(i + NUMBER_FIRST)
                i = i + 1
                df_exl_date = df_exl_filtr_okpo[(df_exl_filtr_okpo['датаРеализации'] == date)].reset_index(drop=True)
                template = os.path.join(dirname, 'maket.docx')
                document = MailMerge(template)
                document.merge(
                    reg_number=df_exl_date['Реєстраційний_номер'][0],
                    doc_tax_number=df_exl_date['Порядковий_№_ПН/РК'][0],
                    doc_tax_date=df_exl_date['Дата_складання_ПН/РК'][0],
                    counterparty_code=df_exl_date['Податковий_номер_Покупця'][0],
                    sum_sale=df_exl_date['Обсяг_операцій'].sum().replace(".", ",")[0],
                    sum_tax=df_exl_date['Сумв_ПДВ'].sum().replace(".", ",")[0],
                    contracte_number=df_exl_date['договорНомер'][0],
                    contracte_date=df_exl_date['договорДата'][0],
                    doc_sale_month=df_exl_date['месяц'].lower()[0],
                    doc_sale_year=df_exl_date['год'][0],
                    doc_sale_number=df_exl_date['номерРеализации'][0],
                    doc_sale_date=df_exl_date['датаРеализации'][0],
                    contracte_count_days=df_exl_date['договорДней'][0],
                    counterpary=df_exl_date['контрагент1С'][0],
                    row=record_number[0],
                    report_date='{:%d.%m.%Y}'.format(datetime.today())
                )

                for doc_type in doc_types:
                    df_pdf_filtered = df_pdf[
                        (df_pdf['датаРеализации'] == date) & (df_pdf['doc_type'] == doc_type)].reset_index(drop=True)
                    doc_number_list = df_pdf_filtered['номерРеализации'].values.tolist().sort()
                    doc_numbers = ', '.join(map(str, doc_number_list))
                    image_save_to_path = ''
                    for i, row in df_pdf_filtered.iterrows():
                        date_revers = datetime.strptime(date, "%d.%m.%Y").strftime("%Y.%m.%d")
                        image_save_to_path = os.path.join(os.path.dirname(row.filename), date_revers, doc_type)
                        print(row.filename)
                        if not os.path.exists(image_save_to_path):  # the folder create here, because we're using row
                            os.makedirs(image_save_to_path)

                        extract_image(row.filename, image_save_to_path)  # extract images from pdf to image_save_to_path
                    doc_number_list = df_merge['номерРеализации'].values.tolist()
                    print(doc_number_list)
                    if os.path.exists(image_save_to_path):
                        # pdf_save_to_path = Path(image_save_to_path,"Group").parents[1]
                        pdf_save_to_path = Path(image_save_to_path, "Group")
                        if not os.path.exists(pdf_save_to_path):  # the folder create here, because we're using row
                            os.makedirs(pdf_save_to_path)
                        add_image_to_pdf(image_save_to_path, pdf_save_to_path)  # add image to pdf

                    # merge_excel_and_word(excel_file_source)

                    print('**************************\n', doc_type, date)

    except Exception as e:
        err_info = "Error: Main: %s" % e
        print(err_info)


if __name__ == '__main__':
    extension = ['*.pdf']
    excel_file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    df_exl, df_pdf = save_df_to_excel(excel_file_source)
    if len(df_exl) > 0 and len(df_pdf) > 0:
        cycle_for_dates(df_exl, df_pdf)
