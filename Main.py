# назначение: в папке отбирает pdf файлы
# создает папки с наименованием = тип документа + дата
# из pdf в эту папку извлекает изображения
# создает 1 pdf файл и все изображения заносит туда
#
import json
import os
import sys

import pandas as pd
from mailmerge import MailMerge
from datetime import datetime
from pathlib import Path
from Image2PdfMultiPages import add_image_to_pdf
from pathvalidate import sanitize_filepath
from PdfExtractImage import extract_image
from MergeExcleWord import edit_excel_and_return_df, convert_date_to_str_df
from Word2Pdf import word_2_pdf

NUMBER_FIRST = 1


# создаем папки по циклу согласно типу_док и дате, и извлекаем туда изображения из pdf
def cycle_for_dates(excel_file_source):
    try:
        df_exl, df_pdf = edit_excel_and_return_df(excel_file_source)
        if len(df_exl) == 0 or len(df_pdf) == 0:
            sys.exit(0)
        dir_name = os.path.dirname(excel_file_source)
        df_pdf = convert_date_to_str_df(df_pdf, 'датаРеализации')
        doc_types = df_pdf['doc_type'].unique()
        client_okpo_list = df_exl['Податковий_номер_Покупця'].unique()
        for client_okpo in client_okpo_list:
            df_exl_okpo = df_exl[(df_exl['Податковий_номер_Покупця'] == client_okpo)].reset_index(drop=True)
            dates = df_exl_okpo['датаРеализации'].unique().tolist()
            client_name = df_exl_okpo['контрагент1С'][0]
            save_to_dir = (os.path.join(dir_name, sanitize_filepath(client_name)))
            if not os.path.isdir(save_to_dir):
                os.mkdir(save_to_dir)
            record_number = NUMBER_FIRST
            for date in dates:
                word_source_df = pd.DataFrame(
                    columns=['doctax_date', 'doctax_number', 'doctax_amount', 'doctax_sumtax', 'reg_number',
                             'counterparty_code', 'total_sale', 'contracte_number', 'contracte_date', 'doc_sale_month',
                             'doc_sale_year', 'contracte_count_days'])
                print(record_number)

                df_exl_okpo_date = df_exl_okpo[(df_exl_okpo['датаРеализации'] == date)].reset_index(drop=True)

                doc_number_list_ttn = df_exl_okpo_date[df_exl_okpo_date['doc_type'].astype(str).str.contains('ТТН')][
                    'номерРеализации'].sort_values().values.tolist()
                doc_numbers_ttn = ', '.join(map(str, doc_number_list_ttn))
                if len(doc_number_list_ttn) > 0:
                    doc_ttn = f"Товаро транспортна накладна № {doc_numbers_ttn} від {date} р."
                else:
                    doc_ttn = ''

                doc_number_list_sale = df_exl_okpo_date[df_exl_okpo_date['doc_type'].astype(str).str.contains('ВН')][
                    'номерРеализации'].sort_values().values.tolist()
                doc_numbers_sale = ', '.join(map(str, doc_number_list_sale))

                for doc_type in doc_types:
                    word_source_doctype = pd.DataFrame(
                        columns=['doctax_date', 'doctax_number', 'doctax_amount', 'doctax_sumtax', 'reg_number',
                                 'counterparty_code', 'total_sale', 'contracte_number', 'contracte_date',
                                 'doc_sale_month', 'doc_sale_year', 'contracte_count_days'])
                    df_pdf_data_doctyoe = df_pdf[
                        (df_pdf['датаРеализации'] == date) & (df_pdf['doc_type'] == doc_type)].reset_index(drop=True)
                    for i, row in df_pdf_data_doctyoe.iterrows():
                        df_exl_okpo_date_doctype = df_exl_okpo_date[
                            df_exl_okpo_date['doc_type'].astype(str).str.contains(row.doc_type)]
                        word_source_doctype['doctax_date'] = df_exl_okpo_date_doctype['Дата_складання_ПН/РК']
                        word_source_doctype['doctax_number'] = df_exl_okpo_date_doctype['Порядковий_№_ПН/РК']
                        word_source_doctype['doctax_amount'] = [x.replace('.', ',') for x in
                                                                df_exl_okpo_date_doctype['Обсяг_операцій'].astype(str)]
                        word_source_doctype['doctax_sumtax'] = [x.replace('.', ',') for x in
                                                                df_exl_okpo_date_doctype['Сумв_ПДВ'].astype(str)]
                        word_source_doctype['reg_number'] = df_exl_okpo_date_doctype['Реєстраційний_номер']
                        word_source_doctype['counterparty_code'] = df_exl_okpo_date_doctype['Податковий_номер_Покупця']
                        word_source_doctype['total_sale'] = df_exl_okpo_date_doctype['Обсяг_операцій']
                        word_source_doctype['contracte_number'] = df_exl_okpo_date_doctype['договорНомер']
                        word_source_doctype['contracte_date'] = df_exl_okpo_date_doctype['договорДата']
                        word_source_doctype['doc_sale_month'] = df_exl_okpo_date_doctype['месяц']
                        word_source_doctype['doc_sale_year'] = df_exl_okpo_date_doctype['год']
                        word_source_doctype['contracte_count_days'] = df_exl_okpo_date_doctype['договорДней']
                        word_source_df = pd.concat([word_source_df, word_source_doctype]).reset_index(drop=True)
                        word_source_df.sort_values(by=['doctax_date', 'doctax_number'], ascending=[True, True],
                                                   inplace=True)
                        date_revers = datetime.strptime(date, "%d.%m.%Y").strftime("%Y.%m.%d")
                        image_save_to_path = os.path.join(save_to_dir, date_revers, doc_type)
                        if not os.path.exists(image_save_to_path):  # the folder create here, because we're using row
                            os.makedirs(image_save_to_path)

                        pdf_save_to_path = Path(image_save_to_path, "Group")
                        if not os.path.exists(pdf_save_to_path):  # the folder create here, because we're using row
                            os.makedirs(pdf_save_to_path)

                        # extract images from pdf to image_save_to_path
                        for i, row in df_pdf_data_doctyoe.iterrows():
                            extract_image(row.filename, image_save_to_path)

                        add_image_to_pdf(image_save_to_path, pdf_save_to_path)  # add image to pdf

                # *********************** source for word
                word_source_df = word_source_df.drop_duplicates()
                json_str = word_source_df.to_json(orient='records')
                # for row in json_str:
                columns = json_str.replace("\\u00a0", "")  # getting rid of empty cells if any there
                columns = json.dumps(columns)
                columns = json.loads(columns)
                array = '{"columns": %s}' % columns
                data = json.loads(array)
                template = os.path.join(dir_name, r'C:\Rasim\Python\ImageToPDF\Maket.docx')
                document = MailMerge(template)
                document.merge_rows('doctax_date', data['columns'])
                document.merge_rows('doctax_number', data['columns'])
                document.merge_rows('doctax_amount', data['columns'])
                document.merge_rows('doctax_sumtax', data['columns'])
                document.merge_rows('reg_number', data['columns'])
                document.merge(
                    counterparty_code=word_source_df['counterparty_code'][0],
                    total_sale=str(round(word_source_df['total_sale'].sum(), 2)).replace(".", ","),
                    contracte_number=word_source_df['contracte_number'][0],
                    contracte_date=word_source_df['contracte_date'][0],
                    doc_sale_month=word_source_df['doc_sale_month'][0],
                    doc_sale_year=word_source_df['doc_sale_year'][0],
                    doc_sale_numbers=doc_numbers_sale,
                    doc_sale_date=date,
                    contracte_count_days=word_source_df['contracte_count_days'][0],
                    counterpary=client_name,
                    docTTN=doc_ttn,
                    row=str(record_number),
                    report_date='{:%d.%m.%Y}'.format(datetime.today())
                )

                record_number = record_number + 1
                word_file = str(Path(os.path.join(save_to_dir, fr"{date}.docx")))
                pdf_file = str(Path(os.path.join(save_to_dir, fr"{date}.pdf")))
                document.write(word_file)  # saving file
                word_2_pdf(word_file, pdf_file)
                print('**************************\n', date)

    except Exception as e:
        err_info = "Error: Main: %s" % e
        print(err_info)

    finally:
        sys.exit(0)


if __name__ == '__main__':
    extension = ['*.pdf']
    excel_file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    cycle_for_dates(excel_file_source)
