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
           "$format=json&$top=10&$inlinecount=allpages&$select=Date,Number"
           f"&$filter=Товары/ДокументОтгрузки eq cast(guid'{invoice_key}',"
           "'Document_РеализацияТоваровУслуг')&$orderby=Date desc")

    return get_response(url)


def get_counterparty_details(date_ttn, number_ttn):
    pass


def get_counterparty_uuid(date_ttn, number_ttn):
    day = parse(date_ttn).day
    month = parse(date_ttn).month
    year = parse(date_ttn).year

    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_скТоварноТранспортнаяНакладная?$format=json"
           f"&$filter=substringof('{number_ttn}', Number) and year(Date) eq {year}"
           f" and month(Date) eq {month} and day(Date) eq {day}")
    return get_response(url)


def add_counterparty_to_ttn_df(df: pd.DataFrame()):
    for i, row in df.iterrows():
        number_ttn = row['1СномерТТН']
        date_ttn = row['1СдатаТТН']


if __name__ == '__main__':
    get_ttn_details('8b108495-507d-11ee-8195-001dd8b72b55')
