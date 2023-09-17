# "ДокументОснование_Type": "StandardODATA.Document_ВозвратТоваровОтПокупателя"
# "ДокументОснование_Type": "StandardODATA.Document_РеализацияТоваровУслуг"
# "ДокументОснование_Type": "StandardODATA.Document_скЛистДоставки"

# find numer TTN by invoice_key
import os
import sys
from dateutil.parser import parse
import pandas as pd
import requests
import json

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"D:\Prestige\Python\Config"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"

sys.path.append(os.path.abspath(CONFIG_PATH))
from configPrestige import DATA_AUTH


def get_response(url):
    result = ''
    try:
        response = requests.get(url, auth=DATA_AUTH)
        if response.status_code == 200:
            result = response.json()['value'][0]
    except:
        print(f"Возникла ошибка при получении данных из url: {url}")

    finally:
        return result


def get_ttn_details(invoice_key):
    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_скТоварноТранспортнаяНакладная?"
           "$format=json"
           f"&$filter=Товары/ДокументОтгрузки eq cast(guid'{invoice_key}','Document_РеализацияТоваровУслуг')"
           "&$select=Ref_Key,Date,Number")

    return get_response(url)


def get_counterparty_name(counterparty_key):
    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Catalog_Контрагенты?$format=json"
           f"&$filter=Ref_Key eq guid'{counterparty_key}'&$select=Description")
    return get_response(url)


def add_ttn_details_to_df(df: pd.DataFrame()):
    df['doc_file_uuid'] = None
    for i, row in df.iterrows():
        date_ttn = row['датаРеализации']
        number_ttn = row['номерРеализации']
        doc_type = row['doc_type']
        if doc_type == 'ВН':
            doc_details = get_doc_sales_details(date_ttn, number_ttn)
        elif doc_type == 'ТТН':
            doc_details = get_doc_transport_details(date_ttn, number_ttn)

        if len(doc_details) == 0:
            print(f"НЕ нашел в 1С документ: {row['filename']}. Перепроверьте наименование скана")
        else:
            doc_ttn_uuid = doc_details['Ref_Key']
            df.loc['doc_file_uuid'] = doc_ttn_uuid

    return df


def get_doc_transport_details(date_ttn, number_ttn):
    day = parse(date_ttn).day
    month = parse(date_ttn).month
    year = parse(date_ttn).year

    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_скТоварноТранспортнаяНакладная?$format=json"
           f"&$filter=substringof('{number_ttn}', Number) and year(Date) eq {year}"
           f" and month(Date) eq {month} and day(Date) eq {day}&$select=Ref_Key,Контрагент,Date,Number")
    return get_response(url)


def get_doc_sales_details(date_ttn, number_ttn):
    date_ttn = str(date_ttn)
    day = parse(date_ttn).day
    month = parse(date_ttn).month
    year = parse(date_ttn).year

    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_РеализацияТоваровУслуг?$format=json"
           f"&$filter=substringof('{number_ttn}', Number) and year(Date) eq {year}"
           f" and month(Date) eq {month} and day(Date) eq {day}&$select=Ref_Key,Контрагент,Date,Number")
    return get_response(url)


def add_to_data_dict(add_to: dict, dict_source: dict) -> dict:
    return add_to.update(dict_source)


if __name__ == '__main__':
    # get_ttn_details('8b108495-507d-11ee-8195-001dd8b72b55')
    doc_transport_details = get_doc_transport_details("05.05.2023", "13606")
    doc_ttn_uuid = doc_transport_details['Ref_Key']
    counterparty_uuid = doc_transport_details['Контрагент']
    counterparty_name = get_counterparty_name(counterparty_uuid)['Description']
    print(counterparty_uuid, counterparty_name)
