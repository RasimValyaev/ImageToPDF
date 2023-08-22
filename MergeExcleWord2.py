# python -m pip install -U pip setuptools
# pip install pandas
# pip install mailmerge
# pip install docx-mailmerge
# pip install xlrd
# pip install xlwt
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b
# https://pbpython.com/python-word-template.html

import pandas
import re
import os
import pandas as pd
from pathvalidate import sanitize_filepath
from datetime import datetime
from mailmerge import MailMerge
from Counterparty import get_counterparty, get_list_of_tax_fatura, get_contract_details, get_doc_sale_details
from Word2Pdf import word_2_pdf
import xlrd
import os.path
from docxtpl import DocxTemplate

MONTH = ['СІЧНЯ', 'ЛЮТОГО', 'БЕРЕЗНЯ', 'КВІТНЯ', 'ТРАВНЯ', 'ЧЕРВНЯ', 'ЛИПНЯ', 'СЕРПНЯ', 'ВЕРЕСНЯ', 'ЖОВТНЯ',
         'ЛИСТОПАДА', 'ГРУДНЯ']


def add_counterparty_name_to_df(path_to_file_excel):
    # added to df counterparty name and code
    df = pandas.read_excel(path_to_file_excel, sheet_name=0)
    df['контрагент1С'] = None
    df['контрагент1Сuuid'] = None

    if df['Дата складання ПН/РК'].dtype == 'int64':
        df["Дата складання ПН/РК"] = df["Дата складання ПН/РК"].map(lambda x: datetime(*xlrd.xldate_as_tuple(x, 0)))

    df["Дата складання ПН/РК"] = pd.to_datetime(df["Дата складання ПН/РК"]).dt.strftime('%d.%m.%Y')

    if df['Дата складання ПН/РК'].dtype == 'int64':
        df["Дата реєстрації ПН/РК в ЄРПН"] = df["Дата реєстрації ПН/РК в ЄРПН"].map(
            lambda x: datetime(*xlrd.xldate_as_tuple(x, 0)))
    df["Дата реєстрації ПН/РК в ЄРПН"] = pd.to_datetime(df["Дата реєстрації ПН/РК в ЄРПН"]).dt.strftime('%d.%m.%Y')

    set_customer_codes = df['Податковий номер Покупця'].unique().tolist()

    for tax_code in set_customer_codes:
        client_uuid, client_name = get_counterparty(tax_code)
        print(client_name)
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1С'] = client_name
        df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1Сuuid'] = client_uuid

    return df


def get_contract(search_doc, list_doc):
    # за день выписано много НН (налогов накл)
    # df содержить их короткие номера.
    # функция по короткому номеру возвращает полный
    for item in list_doc:
        if str(search_doc) in item['Number']:
            return item


def add_doc_tax_details_to_df(df):
    # Search uuid_contracte by date fatura and client_uuid
    df['contract_key'] = None
    df['doc_sale_key'] = None
    df['номерНН_оригинал'] = None
    for i, row in df.iterrows():
        date_doc_tax = row['Дата складання ПН/РК']
        client_uuid = row['контрагент1Сuuid']
        list_of_fatura = get_list_of_tax_fatura(date_doc_tax, client_uuid)
        tax_doc_details = get_contract(row['Порядковий № ПН/РК'], list_of_fatura)
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'номерНН_оригинал'] = tax_doc_details['Number']
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'doc_sale_key'] = tax_doc_details[
            'ДокументОснование']
        df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'contract_key'] = tax_doc_details[
            'ДоговорКонтрагента_Key']

        print(row['Порядковий № ПН/РК'], tax_doc_details)

    return df


def add_doc_contract_details_to_df(df):
    df['договорДней'] = None
    df['договорДата'] = None
    df['договорНомер'] = None
    df['договор'] = None
    set_contract_key = df['contract_key'].unique().tolist()

    for contract_key in set_contract_key:
        contract_details = get_contract_details(contract_key)  # Description,_НКС_ДнівВідтермінуванняОплати,Номер,Дата
        print(contract_details)
        contract_date = datetime.strptime(contract_details['Дата'], "%Y-%m-%dT%H:%M:%S").strftime("%d.%m.%Y")
        df.loc[df['contract_key'] == contract_key, 'договорДней'] = contract_details['_НКС_ДнівВідтермінуванняОплати']
        df.loc[df['contract_key'] == contract_key, 'договорДата'] = contract_date
        df.loc[df['contract_key'] == contract_key, 'договорНомер'] = contract_details['Номер']
        df.loc[df['contract_key'] == contract_key, 'договор'] = contract_details['Description']

    return df


def merge_excel_and_word(path_to_file_excel):
    df = pandas.read_excel(path_to_file_excel, sheet_name='df')
    df = df.astype(str)
    dirname = os.path.dirname(file_source)
    template = os.path.join(dirname, 'maket.docx')

    for i, row in df.iterrows():
        document = MailMerge(template)
        # print(document.get_merge_fields())
        document.merge(
            reg_number=row['Реєстраційний_номер'],
            doc_tax_number=row['Порядковий_№_ПН/РК'],
            doc_tax_date=row['Дата_складання_ПН/РК'],
            counterparty_code=row['Податковий_номер_Покупця'],
            sum_sale=row['Обсяг_операцій'].replace(".", ","),
            sum_tax=row['Сумв_ПДВ'].replace(".", ","),
            contracte_number=row['договорНомер'],
            contracte_date=row['договорДата'],
            doc_sale_month=row['месяц'].lower(),
            doc_sale_year=row['год'],
            doc_sale_number=row['номерРеализации'],
            doc_sale_date=row['датаРеализации'],
            contracte_count_days=row['договорДней'],
            counterpary=row['контрагент1С'],
            row=str(i + 1),
            report_date='{:%d.%m.%Y}'.format(datetime.today())
        )

        save_to_dir = (os.path.join(dirname, sanitize_filepath(row['контрагент1С'])))
        if not os.path.isdir(save_to_dir):
            os.mkdir(save_to_dir)

        # word_file = save_to_dir + fr'/{i + 1}.docx'
        word_file = os.path.join(save_to_dir, fr'{row["filename"]}.docx')
        pdf_file = os.path.join(save_to_dir, fr'{row["filename"]}.pdf')

        document.write(word_file)  # saving file
        word_2_pdf(word_file, pdf_file)
        os.remove(word_file)
        # if i == 4:
        #     break


def add_doc_sale_details_to_df(df):
    df['год'] = None
    df['месяц'] = None
    df['номерРеализации'] = None
    df['датаРеализации'] = None
    for i, row in df.iterrows():
        doc_sale_uuid = row['doc_sale_key']
        doc_sale_details = get_doc_sale_details(doc_sale_uuid)
        doc_sale_month_idx = datetime.strptime(doc_sale_details['Date'], "%Y-%m-%dT%H:%M:%S").month
        doc_sale_month = MONTH[doc_sale_month_idx - 1]
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'номерРеализации'] = int(
            re.findall(r"\d*", doc_sale_details['Number'])[2])
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'датаРеализации'] = doc_sale_details['Date']
        df.loc[df['doc_sale_key'] == row['doc_sale_key'], 'месяц'] = doc_sale_month

    df['год'] = pd.to_datetime(df['датаРеализации']).dt.year
    df['датаРеализации'] = pd.to_datetime(df['датаРеализации']).dt.strftime('%d.%m.%Y')
    return df


def get_valide_columns(df):
    new_columns = []
    for column in df.columns:
        valide_column_name = sanitize_filepath(column)
        valide_column_name = valide_column_name.replace(" ", "_")
        new_columns.append(valide_column_name)

    df.columns = new_columns
    return df


if __name__ == '__main__':
    # file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЛЕГІОН 2015\Написать письмо\Копия ЛЕГІОН 2015.xlsx"
    df = add_counterparty_name_to_df(file_source)
    df = add_doc_tax_details_to_df(df)
    df = add_doc_sale_details_to_df(df)
    df = add_doc_contract_details_to_df(df)
    df = df.drop(columns=['контрагент1Сuuid', 'contract_key', 'doc_sale_key'])
    df['filename'] = df.index + 1
    df.astype(str)
    df['filename'] = pd.concat(["Лист пояснення " + df['filename'].astype(str) + " до " + df[
        r'Дата складання ПН/РК'].astype(str) + " від " + df['датаРеализации'].astype(str)])
    df = get_valide_columns(df)
    with pd.ExcelWriter(file_source, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='df', index=False)

    merge_excel_and_word(file_source)
