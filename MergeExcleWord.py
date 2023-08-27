# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b

import json
import pandas
import pandas as pd
from mailmerge import MailMerge


def merge_excel_and_word(path_to_file_excel):
    excel_data_fragment = pandas.read_excel(path_to_file_excel, sheet_name='source_raw')
    print(excel_data_fragment)
    df = pd.DataFrame()
    df['tax_number'] = excel_data_fragment['Порядковий № ПН/РК'].astype(str)
    df['tax_date'] = excel_data_fragment['Дата складання ПН/РК'].astype(str)

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
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    merge_excel_and_word(file_source)
