# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b
# https://pbpython.com/python-word-template.html

import pandas
import re
import pandas as pd
from datetime import datetime
from mailmerge import MailMerge
from Counterparty import get_counterparty, get_list_of_tax_fatura, get_contract_details, get_doc_sale_details
from Word2Pdf import word_2_pdf

MONTH = ['СІЧНЯ', 'ЛЮТОГО', 'БЕРЕЗНЯ', 'КВІТНЯ', 'ТРАВНЯ', 'ЧЕРВНЯ', 'ЛИПНЯ', 'СЕРПНЯ', 'ВЕРЕСНЯ', 'ЖОВТНЯ',
         'ЛИСТОПАДА', 'ГРУДНЯ']

def add_counterparty_name_to_df(path_to_file_excel):
    # added to df counterparty name and code
    df = pandas.read_excel(path_to_file_excel, sheet_name=0)
    df['контрагент1С'] = None
    df['контрагент1Сuuid'] = None
    df['Дата складання ПН/РК'] = pd.to_datetime(df['Дата складання ПН/РК']).dt.strftime('%d.%m.%Y')
    df['Дата реєстрації ПН/РК в ЄРПН'] = pd.to_datetime(df['Дата реєстрації ПН/РК в ЄРПН']).dt.strftime('%d.%m.%Y')
    set_customer_codes = df['Податковий номер Покупця'].unique().tolist()

    for tax_code in set_customer_codes:
        client_uuid, client_name = get_counterparty(tax_code)
        print(client_name)
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1С'] = client_name
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1Сuuid'] = client_uuid

    return df


def add_doc_tax_details_to_df(df):
    # Search uuid_contracte by date fatura and client_uuid
    df['contract_key'] = None
    df['doc_sale_key'] = None
    df['номерНН_оригинал'] = None
    for i, row in df.iterrows():
        date_fatura = row['Дата складання ПН/РК']
        client_uuid = row['контрагент1Сuuid']
        list_of_fatura = get_list_of_tax_fatura(date_fatura, client_uuid)
        tax_doc = get_contract(row['Порядковий № ПН/РК'], list_of_fatura)
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'номерНН_оригинал'] = tax_doc['Number']
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'doc_sale_key'] = tax_doc[
            'ДокументОснование']
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'contract_key'] = tax_doc[
            'ДоговорКонтрагента_Key']

        print(row['Порядковий № ПН/РК'], tax_doc)

    return df


def merge_excel_and_word(path_to_file_excel):
    dfnew = pd.DataFrame()
    df = pandas.read_excel(path_to_file_excel, sheet_name=0)
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

        document.merge(
            status='Gold',
            city='Springfield',
            phone_number='800-555-5555',
            Business='Cool Shoes',
            zip='55555',
            purchases='$500,000',
            shipping_limit='$500',
            state='MO',
            address='1234 Main Street',
            date='{:%d.%b.%Y}'.format(datetime.today()),
            discount='5%',
            recipient='Mr. Jones')

        template = r'C:\Users\Rasim\Desktop\Scan\tax.docx'
        document = MailMerge(template)

        document.merge_rows('tax_number', data['columns'])
        document.merge_rows('tax_date', data['columns'])
        word_file = fr'C:\Users\Rasim\Desktop\Scan\{i + 1}.docx'
        document.write(word_file)  # saving file

        pdf_file = fr'C:\Users\Rasim\Desktop\Scan\{i + 1}.pdf'
        word_2_pdf(word_file, pdf_file)
        if i == 10:
            break


def get_contract(search_doc, list_doc):
    for item in list_doc:
        if str(search_doc) in item['Number']:
            return item


def add_doc_sale_details_to_df(df):
    df['год'] = None
    df['месяц'] = None
    df['номерРеализации'] = None
    df['датаРеализации'] = None
    for i, row in df.iterrows():
        doc_sale_uuid = row['doc_sale_key']
        doc_sale_details = get_doc_sale_details(doc_sale_uuid)
        doc_sale_month_idx = datetime.strptime(doc_sale_details['Date'],"%Y-%m-%dT%H:%M:%S").month
        doc_sale_month = MONTH[doc_sale_month_idx - 1]
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'номерРеализации'] = int(re.findall(r"\d*",doc_sale_details['Number'])[2])
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'датаРеализации'] = doc_sale_details['Date']
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'месяц'] = doc_sale_month

    df['год'] = pd.to_datetime(df['датаРеализации']).dt.year
    df['датаРеализации'] = pd.to_datetime(df['датаРеализации']).dt.strftime('%d.%m.%Y')
    return df


if __name__ == '__main__':
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    df = add_counterparty_name_to_df(file_source)
    df = add_doc_tax_details_to_df(df)
    df = add_doc_sale_details_to_df(df)
    with pd.ExcelWriter(file_source, mode='a') as writer:
        date = datetime.today().strftime("%d.%m.%Y")
        df.to_excel(writer, sheet_name=date, index=False)

    print(df)
    # merge_excel_and_word(file_source)
