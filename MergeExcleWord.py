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
import warnings
import tkinter as tk
from queue import Empty, Queue
from threading import Thread
from tkinter import filedialog, messagebox  # don't remove. using in Start.py
from dateutil.parser import parse
from pathlib import Path
from pathvalidate import sanitize_filepath
from datetime import datetime
from mailmerge import MailMerge
from DetailsForTaxDocument import get_data_from_taxdoc, get_doc_transport, get_doc_sale_details
from TTN import add_ttn_details_to_df
from Word2Pdf import word_2_pdf
from ConvertXlsToXlsx import convert_xls_to_xlsx

root = tk.Tk()
DATE_PATTERN = r"\b\d{1,2}[,. \-_/\\]\d{1,2}[,. \-_/\\]\d{2,4}\b"

warnings.filterwarnings("ignore", category=UserWarning)

# pd.set_option('precision', 2) # not norking
pd.set_option('float_format', '{:.2f}'.format)

MONTH = ['січні', 'лютому', 'березні', 'квітні', 'травні', 'червні', 'липні', 'серпні', 'вересні', 'жовтні',
         'листопаду', 'грудні']

NUMBER_FIRST = 1  # номер заявления начинается с этого числа


def change_separator(x):
    try:
        x = float(x)
        # Если это удается, то форматируем x с запятой в качестве разделителя
        x = "{:,.2f}".format(x)
    except:
        pass

    return x


def counterparty_name_add_to_df(path_to_file_excel):
    df = pd.DataFrame()
    try:
        df = pd.read_excel(path_to_file_excel, sheet_name=0)
        if not control_columns_name_in_excel_source(df):
            df = pd.DataFrame()
        np.round(df, decimals=2)

        if df['Дата складання ПН/РК'].dtype in ['int64', 'float64']:
            df["Дата складання ПН/РК"] = df["Дата складання ПН/РК"].map(lambda x: datetime(*xlrd.xldate_as_tuple(x, 0)))
        df["Дата складання ПН/РК"] = pd.to_datetime(df["Дата складання ПН/РК"]).dt.strftime('%d.%m.%Y')

        if df["Дата реєстрації ПН/РК в ЄРПН"].dtype in ['int64', 'float64']:
            df["Дата реєстрації ПН/РК в ЄРПН"] = df["Дата реєстрації ПН/РК в ЄРПН"].map(
                lambda x: datetime(*xlrd.xldate_as_tuple(x, 0)))
        df["Дата реєстрації ПН/РК в ЄРПН"] = pd.to_datetime(df["Дата реєстрації ПН/РК в ЄРПН"]).dt.strftime('%d.%m.%Y')

        df['контрагент1С'] = None
        df['counterparty_key'] = None
        df['номерНН_оригинал'] = None
        df['договор'] = None
        df['договорНомер'] = None
        df['договорДата'] = None
        df['договорДней'] = None
        df['ТТН_1Сномер'] = None
        df['ТТН_1Сдата'] = None
        df['НомерТТН_и_ВН_1С'] = None
        df['номерРеализации'] = None
        df['номерРеализацииОриг'] = None
        df['датаРеализации'] = None
        df['месяц'] = None
        df['год'] = None
        df['ТТН_uuid'] = None
        df['invoice_key'] = None
        df['contract_key'] = None
        df['counterparty_key'] = None
        # df[''] = None
        df = df[df['Обсяг операцій'] != 0.00]  # док корректировка(возвраты) не учитывать

        for i, row in df.iterrows():
            try:
                taxdoc = get_data_from_taxdoc(row["Дата складання ПН/РК"], row["Порядковий № ПН/РК"]
                                              , row['ІПН Покупця'])
                if len(taxdoc) > 0:
                    client_uuid = taxdoc['Контрагент']['Ref_Key']
                    df.at[i, 'counterparty_key'] = client_uuid
                    client_name = taxdoc['Контрагент']['Description']
                    taxdoc_number = taxdoc['Number']
                    # print(client_name, taxdoc_number)
                    contract = taxdoc['ДоговорКонтрагента']
                    invoice_uuid = taxdoc['ДокументВводаНаОсновании']
                    if len(invoice_uuid) == 0:
                        invoice_uuid = taxdoc['ДокументОснование']
                    invoice = get_doc_sale_details(invoice_uuid)
                    df.at[i, 'invoice_key'] = invoice_uuid
                    df.at[i, 'contract_key'] = taxdoc['ДоговорКонтрагента_Key']
                    doc_base_uuid = invoice['Сделка']
                    doc_base_type = invoice['Сделка_Type']
                    doc_base_type = doc_base_type.replace('StandardODATA.', '')
                    doc_transport = get_doc_transport(doc_base_uuid, doc_base_type)
                    if doc_transport == '' and doc_base_type == 'Document_ЗаказПокупателя':
                        doc_base_type = 'Document_РеализацияТоваровУслуг'
                        doc_transport = get_doc_transport(invoice_uuid, doc_base_type)

                    if doc_transport != '':
                        df.at[i, 'ТТН_1Сдата'] = parse(doc_transport['Date'], dayfirst=True).strftime("%d.%m.%Y")
                        df.at[i, 'ТТН_1Сномер'] = int(re.search(r"\d+", doc_transport['Number'])[0])
                        df.at[i, 'ТТН_uuid'] = doc_transport['Ref_Key']
                    else:
                        msg = f"в 1С нет ТТН к ВН {invoice['Number']} от {invoice['Date']}"
                        df.at[i, 'ТТН_1Сдата'] = "отсутствует"
                        df.at[i, 'ТТН_1Сномер'] = "отсутствует"
                        print(msg)
                        label = tk.Label(root, text=msg)
                        label.pack()

                    df.at[i, 'контрагент1С'] = client_name
                    df.at[i, 'номерНН_оригинал'] = taxdoc_number
                    df.at[i, 'counterparty_key'] = client_uuid
                    df.at[i, 'договор'] = contract['Description']

                    if contract['Номер'] == '':
                        msg = f"{client_name} Не заполнен номер договора"
                        print(msg)
                        label = tk.Label(root, text=msg)
                        label.pack()
                    else:
                        df.at[i, 'договорНомер'] = contract['Номер']

                    if contract['Дата'] == '0001-01-01T00:00:00':
                        msg = f"{client_name} Не заполнена дата договора"
                        print(msg)
                        label = tk.Label(root, text=msg)
                        label.pack()
                    else:
                        df.at[i, 'договорДата'] = parse(contract['Дата'], dayfirst=True).strftime("%d.%m.%Y")

                    if contract['_НКС_ДнівВідтермінуванняОплати'] in [0, '']:
                        msg = f"{client_name} Не заполнено кол-во дней отсрочки платежа"
                        print(msg)
                        label = tk.Label(root, text=msg)
                        label.pack()
                    else:
                        df.at[i, 'договорДней'] = contract['_НКС_ДнівВідтермінуванняОплати']

                    df.at[i, 'номерРеализацииОриг'] = invoice['Number']
                    df.at[i, 'номерРеализации'] = int(re.search(r"\d+", invoice['Number'])[0])
                    df.at[i, 'датаРеализации'] = parse(invoice['Date'], dayfirst=True).strftime("%d.%m.%Y")
                    df.at[i, 'месяц'] = MONTH[parse(invoice['Date'], dayfirst=True).month - 1]
                    df.at[i, 'год'] = parse(invoice['Date'], dayfirst=True).year
                else:
                    msg = f"Не нашел клиента к НалогНакл {row['Дата складання ПН/РК']} " \
                          f"{row['Порядковий № ПН/РК']} {row['ІПН Покупця']}"
                    print(msg)
                    label = tk.Label(root, text=msg)
                    label.pack()

            except Exception as e:
                msg = str(e)
                print(msg)
                label = tk.Label(root, text=msg)
                label.pack()

    except Exception as e:
        print(str(e))
        pass

    finally:
        return df


def get_valid_columns_name(df):
    new_columns = []
    for column in df.columns:
        valide_column_name = sanitize_filepath(column)
        valide_column_name = valide_column_name.replace(" ", "_")
        new_columns.append(valide_column_name)

    df.columns = new_columns
    return df.reset_index(drop=True)


# convert xls to  xlsx. if input Excel format = xls
def create_new_excel(excel_file):
    filename, file_extension = os.path.splitext(excel_file.lower())
    if file_extension == '.xls':
        excel_file = convert_xls_to_xlsx(excel_file)

    return excel_file


def add_other_parameters_to_df(df):
    try:
        if len(df) > 0:
            today = datetime.today().strftime("%d.%m.%Y")
            df['Лист_пояснення'] = df.index + 1
            df['pdf_filename'] = df.index + NUMBER_FIRST
            df['pdf_filename'] = df['pdf_filename'].apply('{:0>5}'.format)
            df['pdf_filename'] = pd.concat(
                ["Лист пояснення " + df['pdf_filename'].astype(str) + " до " + df[
                    r'Дата складання ПН/РК'].astype(str) + " від " + today])
            df = get_valid_columns_name(df)
            df = df.astype(str)
            df['Обсяг_операцій'] = df['Обсяг_операцій'].astype(float).apply(lambda x: round(x, 2))
            df['Сумв_ПДВ'] = df['Сумв_ПДВ'].astype(float).apply(lambda x: round(x, 2))
            df['датаРеализации'] = pd.to_datetime(df['датаРеализации'], dayfirst=True).dt.strftime('%d.%m.%Y')

        df['НомерТТН_и_ВН_1С'] = np.where(df['номерРеализации'] == df['ТТН_1Сномер'], '', ['Не совпадает'])

    except Exception as e:
        print(str(e))

    finally:
        return df


def counterparty_payment(directory):
    ext = f'{DATE_PATTERN}.(pdf|PDF)$'
    for filename in os.listdir(directory):
        filename = filename.upper()
        try:
            if re.search(r"^(БВ|БB)" + " " + ext, filename):
                print(filename)

        except Exception as e:
            print(str(e))


def fing_incorrect_date(df):
    for i, row in df.iterrows():
        try:
            parse(row['датаРеализации'])
        except:
            msg = "Проверьте на корректность даты в имени файла"
            print(msg, row['filename'])
            label = tk.Label(root, text=msg)
            label.pack()
            continue


# pdf set with date in file name
def get_pdf_set_with_date_in_file_name(excel_path, counterparty_uuid: list):
    doc_type_filter = ['БВ', 'БB', 'РН', 'PH', 'ВН', 'BH', 'TTH', 'ТТН']
    doc_type_ptrn = r"(^[а-яА-ЯёЁa-zA-Z]{2,3})"
    df_all = pd.DataFrame({'filename': os.listdir(excel_path)})
    df_all['filename'] = df_all['filename'].str.upper().reset_index(drop=True)
    df_pdf = df_all[df_all['filename'].str.contains('.PDF')].reset_index(drop=True)
    df = df_pdf[df_pdf['filename'].str.contains('|'.join(doc_type_filter))].reset_index(drop=True)
    df['name'] = df['filename'].str.replace(r'.PDF', '').str.strip().reset_index(drop=True)
    # df['name'] = df['name'].replace(to_replace='[a-zA-Zа-яА-ЯёЁ —]*$', value='', regex=True)
    # df['датаРеализации'] = df['name'].str.extract(DATE_PATTERN, expand=False).str.strip().reset_index(drop=True)
    df["датаРеализации"] = df["name"].apply(lambda x: re.search(DATE_PATTERN, x).group())
    try:
        df['датаРеализации'] = pd.to_datetime(df['датаРеализации'], dayfirst=True)
    except:
        print("В Наименовании файла есть некорректная дата")
        fing_incorrect_date(df)
        sys.exit(0)

    df['doc_type'] = df['name'].str.extract(doc_type_ptrn, expand=False).str.strip().reset_index(drop=True)
    df['номерРеализации'] = df[df['doc_type'] != 'БВ']['name'].str.extract(r"(\d+)", expand=False).str.strip()
    # df.fillna(0, inplace=True)
    df['номерРеализации'] = df['номерРеализации'].astype(pd.Int64Dtype())  # .astype('int64')
    counterparty_date_payment = df[df['doc_type'].str.contains('|'.join(['БВ', 'БB']))][
        'датаРеализации'].sort_values().tolist()
    df_vn_ttn = df[df['doc_type'].str.contains('|'.join(['РН', 'PH', 'ВН', 'BH', 'TTH', 'ТТН']))][
        ['doc_type', 'датаРеализации', 'номерРеализации', 'filename']]
    df_vn_ttn = add_ttn_details_to_df(df_vn_ttn, counterparty_uuid)

    return df_vn_ttn, counterparty_date_payment


def convert_date_to_str_df(df, column_name):
    df[column_name] = df[column_name].astype(str)
    df[column_name] = pd.to_datetime(df[column_name], format="%d.%m.%Y", errors="coerce", dayfirst=True).dt.strftime(
        '%d.%m.%Y')
    return df


# source_from_excel_df = df with data Vika + add other columns
# pdf_df - pdf file names (ВН, ТТН)
def merge_excel_and_word(excel_df, excel_dir, date_of_payments):
    excel_df = excel_df.astype(str)
    template = r"\\PRESTIGEPRODUCT\Scan\Maket.docx"
    print("Ожидайте завершения обработки")
    for i, row in excel_df.iterrows():
        try:
            number_invoice = row['номерРеализации']
            number_transport = row['ТТН_1Сномер']
            if row['файл ВН'] != '':
                doc_sale_header = number_invoice
                number_invoice = f"Видаткова накладна № {number_invoice} від {row['датаРеализации']} р."
            else:
                if number_transport != '':
                    doc_sale_header = number_transport
                else:
                    doc_sale_header = ''

            number_transport = row['ТТН_1Сномер']
            if row['файл ТТН'] != '':
                number_transport = f"Товаро транспортна накладна № {number_transport} від {row['ТТН_1Сдата']} р."

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
                doc_sale_header=doc_sale_header,
                doc_sale_number=doc_sale_header,
                number_invoice=number_invoice,
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
                reg_number=row['Реєстраційний_номер'],
                payments=date_of_payments
            )

            save_to_dir = (os.path.join(excel_dir, sanitize_filepath(row['контрагент1С'])))
            if not os.path.isdir(save_to_dir):
                os.mkdir(save_to_dir)
            word_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.docx")
            pdf_file = os.path.join(save_to_dir, fr"{row['pdf_filename']}.pdf")

            document.write(word_file)  # saving file
            word_2_pdf(word_file, pdf_file)
            print('Создан файл', word_file)
            print('Создан файл', pdf_file)
            # os.remove(word_file)

        except Exception as e:
            print(str(e))


def merge_excel_and_pdf_df(excel_df, pdf_files_df, path_excel):
    if len(pdf_files_df) > 0:  # the sheet "excel_df" need create also if isn't files of pdf
        pdf_files_invoice_df = pdf_files_df[pdf_files_df.doc_type == 'ВН'].reset_index(drop=True)
        pdf_files_invoice_df = pdf_files_invoice_df.rename(columns={'filename': 'файл ВН'})
        pdf_files_invoice_df = pdf_files_invoice_df[['файл ВН', 'doc_file_uuid']]
        pdf_files_transport_df = pdf_files_df[pdf_files_df.doc_type == 'ТТН'].reset_index(drop=True)
        pdf_files_transport_df = pdf_files_transport_df.rename(columns={'filename': 'файл ТТН'})
        pdf_files_transport_df = pdf_files_transport_df[['файл ТТН', 'doc_file_uuid']]
        df_merge = pd.merge(excel_df, pdf_files_invoice_df, how='left', left_on=['invoice_key'],
                            right_on=['doc_file_uuid'])
        df_merge.drop(['doc_file_uuid'], axis=1, inplace=True)  # не удалять!!!
        df_merge = pd.merge(df_merge, pdf_files_transport_df, how='left', left_on=['ТТН_uuid'],
                            right_on=['doc_file_uuid'])
        df_merge.drop(['doc_file_uuid', 'ТТН_uuid', 'invoice_key', 'contract_key', 'counterparty_key'], axis=1,
                      inplace=True)

    else:
        df_merge = excel_df
        df_merge['doc_type'] = None

    # df_merge = df_merge.sort_values("Дата_складання_ПН/РК").reset_index(drop=True)
    df_merge = convert_date_to_str_df(df_merge, 'датаРеализации')
    df_merge = convert_date_to_str_df(df_merge, 'Дата_складання_ПН/РК')
    df_merge = convert_date_to_str_df(df_merge, 'Дата_реєстрації_ПН/РК_в_ЄРПН')
    df_merge = convert_date_to_str_df(df_merge, 'договорДата')
    df_merge.fillna('')
    save_as = Path(Path(path_excel).parent, Path(path_excel).stem + '_new' + Path(path_excel).suffix)
    # with pd.ExcelWriter(save_as, mode='a', if_sheet_exists='new') as writer:
    #     df_merge.to_excel(writer, sheet_name='excel_df', index=False)

    df_merge['договорНомер'] = df_merge['договорНомер'].fillna("отсутствует")
    df_merge['договорДата'] = df_merge['договорДата'].fillna("отсутствует")
    df_merge["Статус_ПН/РК"] = df_merge["Статус_ПН/РК"].apply(change_separator)
    try:
        with pd.ExcelWriter(save_as) as writer:
            df_merge.to_excel(writer, sheet_name='excel_df', index=False)
        msg = "Создан файл %s" % save_as
        print(msg)
        label = tk.Label(root, text=msg)
        label.pack()

    except Exception as e:
        msg = f"Ошибка при сохранении файла в Excel: {str(e)}.\nВозможно файл {save_as} уже открыт"
        print(msg)
        label = tk.Label(root, text=msg)
        label.pack()

    finally:
        return df_merge


def control_columns_name_in_excel_source(df):
    need_columns = ['Реєстраційний номер', 'Дата складання ПН/РК', 'Дата реєстрації ПН/РК в ЄРПН',
                    'Податковий номер Покупця', 'Обсяг операцій', 'Сумв ПДВ']
    all_in_df = all(item in df.columns for item in need_columns)
    if not all_in_df:
        print("ERROR: Проверьте, чтобы Excel содержал все колонки", need_columns)

    return all_in_df


# create sheet "df" in current file_excel
# if there are pdf files on current folder added doc_type in Excel
def excel_to_df(excel_file):
    try:
        # excel_file = create_new_excel(excel_file)
        df = counterparty_name_add_to_df(excel_file)
        if len(df) != 0:
            df = add_other_parameters_to_df(df)

    except Exception as e:
        err_info = "Error: MergeExcleWord: %s" % e
        if e.args[0] == 13:
            print("Закройте файл {}".format(excel_file))
            sys.exit(0)
        else:
            print(err_info)

    finally:
        return df.astype(str)


def get_bank_statement(date_of_payments):
    result = ''
    size = len(date_of_payments)
    if size == 0:
        msg = ("\n*****************************************************************"
               "\nВыписки банка в формате pdf в текущем каталоге НЕ обнаружены."
               "\nФормирование писем ПРЕКРАЩЕНО"
               "\n*****************************************************************\n"
               )
        print(msg)
        label = tk.Label(root, text=msg)
        label.pack()

    else:
        for i, date in enumerate(date_of_payments):
            if i != (size - 1):
                result += f"{i + 4}. Банківська виписка від {'{:%d.%m.%Y}'.format(date)}р.\n"
            else:
                result += f"{i + 4}. Банківська виписка від {'{:%d.%m.%Y}'.format(date)}р."
    return result


def get_counterparty_uuid_from_excel_df(df: pd.DataFrame()):
    counterparty_uuid = df.groupby('counterparty_key')
    return counterparty_uuid.counterparty_key.unique().to_list()


def merge_excle_word_main(excel_file, template):
    if not os.path.exists(excel_file):
        msg = ("Не найден Excel файл", excel_file)
        print(msg)
        label = tk.Label(root, text=msg)
        label.pack()

        sys.exit(0)

    # creating df with pdf files
    excel_dir = Path(os.path.dirname(excel_file))

    # add other column to source - excel + adding df with pdf files
    excel_df = excel_to_df(excel_file)
    if len(excel_df) != 0:
        counterparty_uuid = get_counterparty_uuid_from_excel_df(excel_df)

        pdf_files_df, date_of_payment = get_pdf_set_with_date_in_file_name(excel_dir, counterparty_uuid)
        date_of_payments = get_bank_statement(date_of_payment)

        if len(pdf_files_df) == 0:
            msg = "Не обнаружены сканы начинающиеся на ВН, ТТН, БВ\n"
            print(msg)
            label = tk.Label(root, text=msg)
            label.pack()

        df_merge = merge_excel_and_pdf_df(excel_df, pdf_files_df, excel_file)
        if date_of_payments != '' and template.exists():
            merge_excel_and_word(df_merge, excel_dir, date_of_payments)


if __name__ == '__main__':
    excel_file_source = r"\\PRESTIGEPRODUCT\Scan\Левайс\Левайс.xls"
    # excel_file_source = r"c:\Users\Rasim\Desktop\на суд с колич.xls"
    merge_excle_word_main(excel_file_source)
