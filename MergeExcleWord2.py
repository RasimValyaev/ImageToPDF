# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b

import json
import pandas
import pandas as pd
import win32com.client as win32

from mailmerge import MailMerge

from Word2Pdf import word_2_pdf


def merge_excel_and_word(path_to_file_excel):
    df_source = pandas.read_excel(path_to_file_excel, sheet_name='source_raw')
    # df_source['Порядковий № ПН/РК'].astype(str)
    # df_source['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    print(df_source)
    df = pd.DataFrame()
    df['tax_number'] = df_source['Порядковий № ПН/РК'].astype(str)
    df['tax_date'] = df_source['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    for i, row in df.iterrows():
        df_list = []
        data = {}
        df_list.append(row.to_dict())  # row.to_json(orient='records')
        data['columns'] = df_list
        data = json.dumps(data)
        data = json.loads(data)
        template = r'C:\Users\Rasim\Desktop\Scan\tax.docx'
        document = MailMerge(template)

        document.merge_rows('tax_number', data['columns'])
        document.merge_rows('tax_date', data['columns'])
        word_file = fr'C:\Users\Rasim\Desktop\Scan\{i+1}.docx'
        document.write(word_file)  # saving file

        pdf_file = fr'C:\Users\Rasim\Desktop\Scan\{i+1}.pdf'
        word_2_pdf(word_file, pdf_file)
        if i == 10:
            break


if __name__ == '__main__':
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    merge_excel_and_word(file_source)
