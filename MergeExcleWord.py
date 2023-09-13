# python -m pip install -U pip setuptools
# pip install openpyxl
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
import sys
import numpy as np
import pandas as pd
import os.path
import xlrd
from pathlib import Path
from pathvalidate import sanitize_filepath
from datetime import datetime
from mailmerge import MailMerge
from DetailsForTaxDocument import get_counterparty, get_list_of_tax_fatura, get_contract_details, get_doc_sale_details
from Word2Pdf import word_2_pdf
from ConvertXlsToXlsx import convert_xls_to_xlsx

# pd.set_option('precision', 2) # not norking
pd.set_option('float_format', '{:.2f}'.format)

MONTH = ['січні', 'лютому', 'березні', 'квітні', 'травні', 'червні', 'липні', 'серпні', 'вересні', 'жовтні',
         'листопаду', 'грудні']

NUMBER_FIRST = 1  # номер заявления начинается с этого числа


def counterparty_name_add_to_df(path_to_file_excel):
    df = pd.DataFrame()
    # added to df counterparty name and code
    try:
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
            if tax_code == 0:
                print('stop')
            try:
                counterparty = get_counterparty(tax_code)
                if len(counterparty) > 1:
                    client_uuid, client_name = counterparty
                    print(client_name)
                    df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1С'] = client_name
                    df.loc[df['Податковий номер Покупця'] == tax_code, 'контрагент1Сuuid'] = client_uuid
            except Exception as e:
                print(str(e))

    except Exception as e:
        print(str(e))

    finally:
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
        try:
            date_doc_tax = row['Дата складання ПН/РК']
            client_uuid = row['контрагент1Сuuid']
            list_of_fatura = get_list_of_tax_fatura(date_doc_tax, client_uuid)
            tax_doc_details = get_contract(row['Порядковий № ПН/РК'], list_of_fatura)
            df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'номерНН_оригинал'] = tax_doc_details[
                'Number']
            df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'doc_sale_key'] = tax_doc_details[
                'ДокументОснование']
            df.loc[df['Порядковий № ПН/РК'] == row['Порядковий № ПН/РК'], 'contract_key'] = tax_doc_details[
                'ДоговорКонтрагента_Key']

        except Exception as e:
            print(str(e))
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


def get_doctype_by_docnumber(pdf_file_names: pd.DataFrame(), look_number: str, doc_type: str):
    doc_sale = pdf_file_names[(pdf_file_names['doc_type'] == doc_type)
                              & (pdf_file_names['номерРеализации'] == look_number)]
    if len(doc_sale) > 0:
        return look_number
    else:
        return ''


# source_from_excel_df = df with data Vika + add other columns
# pdf_df - pdf file names (ВН, ТТН)
def merge_excel_and_word(source_from_excel_df, pdf_df):
    # df = pd.read_excel(excel_file_source, sheet_name='df')
    source_from_excel_df = source_from_excel_df.astype(str)
    dirname = os.path.dirname(excel_file_source)
    template = r"C:\Rasim\Python\ImageToPDF\Maket.docx"

    for i, row in source_from_excel_df.iterrows():
        doc_number = row['номерРеализации']
        number_invoice = get_doctype_by_docnumber(pdf_df, doc_number, 'ВН')
        number_transport = get_doctype_by_docnumber(pdf_df, doc_number, 'ТТН')

        if number_transport != '':
            number_transport = f"Товаро транспортна накладна № {number_transport} від {row['датаРеализации']} р."

        if number_invoice == '' and number_transport == '':
            continue

        record_number = str(row['Лист_пояснення'])
        document = MailMerge(template)
        document.merge(
            counterparty_code=row['Податковий_номер_Покупця'],
            total_sale=row['Обсяг_операцій'].replace(".", ","),
            contracte_number=row['договорНомер'],
            contracte_date=row['договорДата'],
            doc_sale_month=row['месяц'],
            doc_sale_year=row['год'],
            doc_sale_numbers=number_invoice,
            doc_sale_date=row['датаРеализации'],
            contracte_count_days=row['договорДней'],
            counterpary=row['контрагент1С'],
            docTTN=number_transport,
            row=record_number,
            report_date='{:%d.%m.%Y}'.format(datetime.today()),
            doctax_date=row['Дата_складання_ПН/РК'],
            doctax_number=row['Порядковий_№_ПН/РК'],
            doctax_amount=row['Обсяг_операцій'].replace(".", ","),
            doctax_sumtax=row['Сумв_ПДВ'].replace(".", ","),
            reg_number=row['Реєстраційний_номер']
        )

        save_to_dir = (os.path.join(dirname, sanitize_filepath(row['контрагент1С'])))
        if not os.path.isdir(save_to_dir):
            os.mkdir(save_to_dir)
        word_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.docx")
        pdf_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.pdf")

        document.write(word_file)  # saving file
        word_2_pdf(word_file, pdf_file)
        # os.remove(word_file)


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


def add_other_parameters_to_df(excel_file):
    df = pd.DataFrame()
    try:
        filename, file_extension = os.path.splitext(excel_file.lower())
        if file_extension == '.xls':
            excel_file = convert_xls_to_xlsx(excel_file)

        df = counterparty_name_add_to_df(excel_file)
        if len(df) > 0:
            df = doc_tax_details_add_to_df(df)
            if len(df) > 0:
                df = doc_sale_details_add_to_df(df)
                if len(df) > 0:
                    df = doc_contract_details_add_to_df(df)
                    if len(df) > 0:
                        today = datetime.today().strftime("%d.%m.%Y")
                        df = df.drop(columns=['контрагент1Сuuid', 'contract_key', 'doc_sale_key'])
                        df['Лист_пояснення'] = df.index + 1
                        df['pdf_filename'] = df.index + NUMBER_FIRST
                        # df['pdf_filename'] = df['pdf_filename'].apply('{:0>5}'.format)
                        df.astype(str)
                        df['pdf_filename'] = pd.concat(
                            ["Лист пояснення " + df['pdf_filename'].astype(str) + " до " + df[
                                r'Дата складання ПН/РК'].astype(str) + " від " + today])
                        df = get_validcolumns_name(df)
                        df = df.astype(str)
                        try:
                            df['Статус_ПН/РК'] = df['Статус_ПН/РК'].astype(float).apply(lambda x: round(x, 2))
                        except Exception as e:
                            print(str(e))
                        df['Обсяг_операцій'] = df['Обсяг_операцій'].astype(float).apply(lambda x: round(x, 2))
                        df['Сумв_ПДВ'] = df['Сумв_ПДВ'].astype(float).apply(lambda x: round(x, 2))
                        df['номерРеализации'] = df['номерРеализации'].apply(int)
                        df['датаРеализации'] = df['датаРеализации'].apply(pd.to_datetime, format='%d.%m.%Y')

    except Exception as e:
        print(str(e))

    finally:
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
        try:
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

        except Exception as e:
            print(str(e))

    df = pd.DataFrame(data)
    if len(df) > 0:
        df['датаРеализации'] = df['датаРеализации'].apply(pd.to_datetime, format='%d.%m.%Y')

    return df


def convert_date_to_str_df(df, column_name):
    if df[column_name].dtype == '<M8[ns]':
        df[column_name] = df[column_name].dt.strftime('%d.%m.%Y')
    else:
        df[column_name] = pd.to_datetime(df[column_name], format='%d.%m.%Y').dt.strftime('%d.%m.%Y')
    return df


# create sheet "df" in current file_excel
# if there are pdf files on current folder added doc_type in Excel
def edit_excel_and_return_df(excel_file):
    df_merge = pd.DataFrame()
    pdf_files_df = pd.DataFrame()
    try:
        dir_name = os.path.dirname(excel_file)
        df = add_other_parameters_to_df(excel_file)
        pdf_files_df = get_pdf_set_with_date_in_file_name(
            dir_name)  # dataframe with pdf filenames from folder with excel_file_source
        if len(pdf_files_df) > 0:  # the sheet "df" need create also if isn't files of pdf
            type_of_docs_df = pdf_files_df.groupby(['датаРеализации', 'номерРеализации'],
                                                   as_index=False)['doc_type'].agg(list)
            df_merge = pd.merge(df, type_of_docs_df, how='left', left_on=['датаРеализации', 'номерРеализации'],
                                right_on=['датаРеализации', 'номерРеализации'])
        else:
            df_merge = df
            df_merge['doc_type'] = None

        df_merge = df_merge.sort_values("датаРеализации")
        df_merge = convert_date_to_str_df(df_merge, 'датаРеализации')
        df_merge = convert_date_to_str_df(df_merge, 'Дата_складання_ПН/РК')
        df_merge = convert_date_to_str_df(df_merge, 'Дата_реєстрації_ПН/РК_в_ЄРПН')
        df_merge = convert_date_to_str_df(df_merge, 'договорДата')
        df_merge.rename(columns={'doc_type': 'doc_type_list'})
        df_merge.fillna('')
        save_as = Path(Path(excel_file).parent, Path(excel_file).stem + '_new' + Path(excel_file).suffix)
        # with pd.ExcelWriter(save_as, mode='a', if_sheet_exists='new') as writer:
        #     df_merge.to_excel(writer, sheet_name='df', index=False)
        with pd.ExcelWriter(save_as) as writer:
            df_merge.to_excel(writer, sheet_name='df', index=False)

    except Exception as e:
        err_info = "Error: MergeExcleWord: %s" % e
        if e.args[0] == 13:
            print("Закройте файл {}".format(excel_file))
            sys.exit(0)
        else:
            print(err_info)

    finally:
        return df_merge.astype(str), pdf_files_df.astype(str)


if __name__ == '__main__':
    excel_file_source = r"c:\Users\Rasim\Desktop\Scan\ДЕЛІКАТ РИТЕЙЛ\ДЕЛІКАТ РИТЕЙЛ.xlsx"
    df_merge, pdf_files_df = edit_excel_and_return_df(excel_file_source)
    # print(df_merge)
    # print(pdf_files_df)
    merge_excel_and_word(df_merge, pdf_files_df)
