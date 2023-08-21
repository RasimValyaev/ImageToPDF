# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b

import json
import pandas
import pandas as pd
from datetime import datetime
import win32com.client as win32

from mailmerge import MailMerge

from Counterparty import get_counterparty, get_list_of_fatura
from Word2Pdf import word_2_pdf


def add_counterparty_name(path_to_file_excel):
    df = pandas.read_excel(path_to_file_excel, sheet_name=0)
    df['контрагент1С'] = None
    df['контрагент1Сuuid'] = None
    df['Дата складання ПН/РК'] = df['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    df['Дата реєстрації ПН/РК в ЄРПН'] = df['Дата реєстрації ПН/РК в ЄРПН'].dt.strftime('%d.%m.%Y')
    unq = df['Податковий номер Покупця'].unique().tolist()

    for tax_code in unq:
        client_uuid, client_name = get_counterparty(tax_code)
        print(client_name)
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1С'] = client_name
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1Сuuid'] = client_uuid

    return df  # df.astype(str)


def merge_excel_and_word(path_to_file_excel):
    dfnew = pd.DataFrame()
    df = pandas.read_excel(path_to_file_excel, sheet_name='source_raw')
    df['counterparty'] = None
    df['Дата складання ПН/РК'] = df['Дата складання ПН/РК'].dt.strftime('%d.%m.%Y')
    df['Дата реєстрації ПН/РК в ЄРПН'] = df['Дата реєстрації ПН/РК в ЄРПН'].dt.strftime('%d.%m.%Y')
    for i, row in df.iterrows():
        df_list = []
        data = {}
        tax_code = row['ІПН Покупця']
        client_name = get_counterparty(tax_code)
        print(client_name)
        dfnew.loc[df['ІПН Покупця'] == row['ІПН Покупця'], 'counterparty'] = client_name
        df_list.append(row.to_dict())  # row.to_json(orient='records')
        data['columns'] = df_list
        data = json.dumps(data)
        data = json.loads(data)
        template = r'C:\Users\Rasim\Desktop\Scan\tax.docx'
        document = MailMerge(template)

        document.merge_rows('tax_number', data['columns'])
        document.merge_rows('tax_date', data['columns'])
        word_file = fr'C:\Users\Rasim\Desktop\Scan\{i + 1}.docx'
        document.write(word_file)  # saving file

        pdf_file = fr'C:\Users\Rasim\Desktop\Scan\{i + 1}.pdf'
        word_2_pdf(word_file, pdf_file)
        # if i == 10:
        #     break


def get_contract(search_doc, list_doc):
    for item in list_doc:
        if str(search_doc) in item['Number']:
            return item.values()
        else:
            continue


def add_contract(df):
    df['договор'] = None
    for i, row in df.iterrows():
        date_fatura = datetime.strptime(row['Дата складання ПН/РК'], '%d.%m.%Y')
        client_uuid = row['контрагент1Сuuid']
        list_of_fatura = get_list_of_fatura(date_fatura, client_uuid)
        tax_doc = get_contract(row['Порядковий № ПН/РК'], list_of_fatura)
        print(row['Порядковий № ПН/РК'], tax_doc)
        pass
    # print(list_of_fatura)


if __name__ == '__main__':
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    df = add_counterparty_name(file_source)
    add_contract(df)
    with pd.ExcelWriter(file_source, mode='a') as writer:
        date = datetime.today().strftime("%d.%m.%Y")
        df.to_excel(writer, sheet_name=date, index=False)

    print(df)
    # merge_excel_and_word(file_source)
