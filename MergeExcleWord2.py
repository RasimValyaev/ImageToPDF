# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b

import json

import pandas
import pandas as pd
from mailmerge import MailMerge


def merge_excel_and_word(path_to_file_excel):
    df_source = pandas.read_excel(path_to_file_excel, sheet_name='source_raw')
    # df_source['Порядковий № ПН/РК'].astype(str)
    # df_source['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    print(df_source)
    df = pd.DataFrame()
    df['tax_number'] = df_source['Порядковий № ПН/РК'].astype(str)
    df['tax_date'] = df_source['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    data = {}
    for i, row in df.iterrows():
        data['columns'] = row.to_dict()  # row.to_json(orient='records')
        data = json.dumps(data)
        template = r'C:\Users\Rasim\Desktop\Scan\tax.docx'
        document = MailMerge(template)

        document.merge_rows('tax_number', data['columns'])
        document.merge_rows('tax_date', data['columns'])
        document.write(r'C:\Users\Rasim\Desktop\Scan\alpha-output.docx')  # saving file
        break


if __name__ == '__main__':
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    merge_excel_and_word(file_source)
