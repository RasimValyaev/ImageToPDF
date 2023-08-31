# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b

import json
import pandas
from mailmerge import MailMerge


def merge_excel_and_word(df):
    # df['tax_number'] = df['Порядковий № ПН/РК'].astype(str)
    # df['tax_date'] = df['Дата складання ПН/РК'].astype(str)
    df['doc_tax_date'] = df['Дата складання ПН/РК'].astype(str)
    df['doc_tax_number'] = df['Порядковий_№_ПН/РК'].astype(str)
    df['sum_sale'] = df['Обсяг_операцій'].astype(str)
    df['reg_number'] = df['Реєстраційний_номер'].astype(str)

    json_str = df.to_json(orient='records')
    # for row in json_str:
    columns = json_str.replace("\\u00a0", "")  # getting rid of empty cells if any there
    columns = json.dumps(columns)
    columns = json.loads(columns)
    array = '{"columns": %s}' % columns
    data = json.loads(array)

    template = r'C:\Users\Rasim\Desktop\Scan\tax.docx'
    document = MailMerge(template)

    document.merge_rows('tax_number', data['columns'])
    document.merge_rows('tax_date', data['columns'])
    document.write(r'C:\Users\Rasim\Desktop\Scan\alpha-output.docx')  # saving file


if __name__ == '__main__':
    path_to_file_excel = r"c:\Users\Rasim\Desktop\Scan\РелайзКомпани\Релайз для разблокировки.xlsx"
    df = pandas.read_excel(path_to_file_excel, sheet_name=0)
    merge_excel_and_word(df)
