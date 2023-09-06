# python -m pip install -U pip setuptools
# pip install pandas
# pip install mailmerge - не нужен
# pip install docx-mailmerge
# pip install xlrd
# pip install pathvalidate
# pip install xlwt - не нужен
# https://archit-narain.medium.com/how-to-merge-tables-to-word-documents-using-python-9786124a276b
# https://pbpython.com/python-word-template.html

# берет данные из Excel и подставляет их в шаблон Word, создает новый док Word

import re
import os
import numpy as np
import pandas as pd
import os.path
import xlrd
from pathvalidate import sanitize_filepath
from datetime import datetime
from mailmerge import MailMerge
from DetailsForTaxDocument import get_counterparty, get_list_of_tax_fatura, get_contract_details, get_doc_sale_details
from Word2Pdf import word_2_pdf
from ConvertXlsToXlsx import convert_xls_to_xlsx

# pd.set_option('precision', 2)
pd.set_option('float_format', '{:.2f}'.format)

MONTH = ['січні', 'лютому', 'березні', 'квітні', 'травні', 'червні', 'липні', 'серпні', 'вересні', 'жовтні',
         'листопаду', 'грудні']

NUMBER_FIRST = 149 + 1  # номер заявления начинается с этого числа


def counterparty_name_add_to_df(path_to_file_excel):
    # added to df counterparty name and code
    df = pd.read_excel(path_to_file_excel, sheet_name=0)
    np.round(df, decimals=2)
    df['контрагент1С'] = None
    df['контрагент1Сuuid'] = None

    if df['Дата складання ПН/РК'].dtype in ['int64', 'float64']:
        df["Дата складання ПН/РК"] = df["Дата складання ПН/РК"].map(lambda x: datetime(*xlrd.xldate_as_tuple(x, 0)))
    df["Дата складання ПН/РК"] = pd.to_datetime(df["Дата складання ПН/РК"]).dt.strftime('%d.%m.%Y')

    if df["Дата реєстрації ПН/РК в ЄРПН"].dtype in ['int64', 'float64']:
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


def doc_tax_details_add_to_df(df):
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

    return df


def doc_contract_details_add_to_df(df):
    df['договорДней'] = None
    df['договорДата'] = None
    df['договорНомер'] = None
    df['договор'] = None
    set_contract_key = df['contract_key'].unique().tolist()

    for contract_key in set_contract_key:
        contract_details = get_contract_details(contract_key)  # Description,_НКС_ДнівВідтермінуванняОплати,Номер,Дата
        contract_date = datetime.strptime(contract_details['Дата'], "%Y-%m-%dT%H:%M:%S").strftime("%d.%m.%Y")
        df.loc[df['contract_key'] == contract_key, 'договорДней'] = contract_details['_НКС_ДнівВідтермінуванняОплати']
        df.loc[df['contract_key'] == contract_key, 'договорДата'] = contract_date
        df.loc[df['contract_key'] == contract_key, 'договорНомер'] = contract_details['Номер']
        df.loc[df['contract_key'] == contract_key, 'договор'] = contract_details['Description']

    return df


def merge_excel_and_word(path_to_file_excel):
    df = pd.read_excel(path_to_file_excel, sheet_name='df')
    df = df.astype(str)
    dirname = os.path.dirname(excel_file_source)
    template = os.path.join(dirname, 'maket.docx')

    for i, row in df.iterrows():
        record_number = str(i + NUMBER_FIRST)
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
            row=record_number,
            report_date='{:%d.%m.%Y}'.format(datetime.today())
        )

        save_to_dir = (os.path.join(dirname, sanitize_filepath(row['контрагент1С'])))
        if not os.path.isdir(save_to_dir):
            os.mkdir(save_to_dir)

        # word_file = save_to_dir + fr'/{i + 1}.docx'
        word_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.docx")
        pdf_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.pdf")

        document.write(word_file)  # saving file
        word_2_pdf(word_file, pdf_file)
        os.remove(word_file)


def doc_sale_details_add_to_df(df):
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


def get_validcolumns_name(df):
    new_columns = []
    for column in df.columns:
        valide_column_name = sanitize_filepath(column)
        valide_column_name = valide_column_name.replace(" ", "_")
        new_columns.append(valide_column_name)

    df.columns = new_columns
    return df


def add_other_parameters_to_df(excel_file_source):
    filename, file_extension = os.path.splitext(excel_file_source.lower())
    if file_extension == '.xls':
        excel_file_source = convert_xls_to_xlsx(excel_file_source)

    df = counterparty_name_add_to_df(excel_file_source)
    if len(df) > 0:
        df = doc_tax_details_add_to_df(df)
        if len(df) > 0:
            df = doc_sale_details_add_to_df(df)
            if len(df) > 0:
                df = doc_contract_details_add_to_df(df)
                if len(df) > 0:
                    df = df.drop(columns=['контрагент1Сuuid', 'contract_key', 'doc_sale_key'])
                    df['pdf_filename'] = df.index + NUMBER_FIRST
                    df['pdf_filename'] = df['pdf_filename'].apply('{:0>5}'.format)
                    df.astype(str)
                    df['pdf_filename'] = pd.concat(["Лист пояснення " + df['pdf_filename'].astype(str) + " до " + df[
                        r'Дата складання ПН/РК'].astype(str) + " від " + df['датаРеализации'].astype(str)])
                    df = get_validcolumns_name(df)
                    df = df.astype(str)
                    df['Статус_ПН/РК'] = df['Статус_ПН/РК'].astype(float).apply(lambda x: round(x, 2))
                    df['Обсяг_операцій'] = df['Обсяг_операцій'].astype(float).apply(lambda x: round(x, 2))
                    df['Сумв_ПДВ'] = df['Сумв_ПДВ'].astype(float).apply(lambda x: round(x, 2))
                    df['номерРеализации'] = df['номерРеализации'].apply(int)
                    df['датаРеализации'] = df['датаРеализации'].apply(pd.to_datetime, format='%d.%m.%Y')

    return df


# pdf set with date in file name
def get_pdf_set_with_date_in_file_name(directory):
    ext = r'\d{2}.\d{2}.\d{4}.pdf$'
    data = {}
    doc_type_list = []
    date_list = []
    file_list = []
    doc_number_list = []
    for filename in os.listdir(directory):
        if re.search(ext, filename):
            file_list.append(os.path.join(directory, filename))
            doc_type_list.append(re.search('[а-яА-ЯёЁa-zA-Z]+', filename)[0])
            date_raw = re.search("\d{1,22}[.,]\d{1,2}[.,]\d{2,4}", filename)
            if date_raw:
                date = date_raw[0].replace(",", ".")
                date_list.append(date)
            else:
                continue

            doc_number = re.search(" \d+ ", filename)
            if doc_number:
                doc_number_list.append(int(re.search(" \d+ ", filename)[0]))
            else:
                doc_number_list.append(None)

            data.update(
                {"doc_type": doc_type_list, "датаРеализации": date_list, "номерРеализации": doc_number_list,
                 "filename": file_list})

    df = pd.DataFrame(data)
    df['датаРеализации'] = df['датаРеализации'].apply(pd.to_datetime, format='%d.%m.%Y')
    return df


def convert_date_to_str_df(df, column_name):
    if df[column_name].dtype == '<M8[ns]':
        df[column_name] = df[column_name].dt.strftime('%d.%m.%Y')
    else:
        df[column_name] = pd.to_datetime(df[column_name], format='%d.%m.%Y').dt.strftime('%d.%m.%Y')
    return df


def edit_excel_and_return_df(excel_file_source):
    df_merge = pd.DataFrame()
    pdf_files_df = pd.DataFrame()
    try:
        dir_name = os.path.dirname(excel_file_source)
        df = add_other_parameters_to_df(excel_file_source)
        pdf_files_df = get_pdf_set_with_date_in_file_name(
            dir_name)  # dataframe with pdf filenames from folder with excel_file_source
        type_of_docs_df = pdf_files_df.groupby(['датаРеализации', 'номерРеализации'],
                                               as_index=False)['doc_type'].agg(list)
        df_merge = pd.merge(df, type_of_docs_df, how='left', left_on=['датаРеализации', 'номерРеализации'],
                            right_on=['датаРеализации', 'номерРеализации'])
        df_merge = df_merge.sort_values("датаРеализации")
        df_merge = convert_date_to_str_df(df_merge, 'датаРеализации')
        df_merge = convert_date_to_str_df(df_merge, 'Дата_складання_ПН/РК')
        df_merge = convert_date_to_str_df(df_merge, 'Дата_реєстрації_ПН/РК_в_ЄРПН')
        df_merge = convert_date_to_str_df(df_merge, 'договорДата')
        df_merge.rename(columns={'doc_type': 'doc_type_list'})
        df_merge.fillna('')
        with pd.ExcelWriter(excel_file_source, mode='a', if_sheet_exists='replace') as writer:
            df_merge.to_excel(writer, sheet_name='df', index=False)

    except Exception as e:
        err_info = "Error: MergeExcleWord2: %s" % e
        print(err_info)

    finally:
        return df_merge, pdf_files_df


if __name__ == '__main__':
    excel_file_source = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЄВРО СМАРТ ПАУЕР\ТОВ ЄВРО СМАРТ ПАУЕР.xlsx"
    edit_excel_and_return_df(excel_file_source)
