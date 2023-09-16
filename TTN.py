# "ДокументОснование_Type": "StandardODATA.Document_ВозвратТоваровОтПокупателя"
# "ДокументОснование_Type": "StandardODATA.Document_РеализацияТоваровУслуг"
# "ДокументОснование_Type": "StandardODATA.Document_скЛистДоставки"

# find numer TTN by invoice_key
import os
import sys
import requests
import json

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"D:\Prestige\Python\Config"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"

sys.path.append(os.path.abspath(CONFIG_PATH))
from configPrestige import DATA_AUTH


def get_ttn_details(invoice_key):
    result = ''
    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_скТоварноТранспортнаяНакладная?"
           "$format=json&$top=10&$inlinecount=allpages&$select=Date,Number"
           f"&$filter=Товары/ДокументОтгрузки eq cast(guid'{invoice_key}',"
           "'Document_РеализацияТоваровУслуг')&$orderby=Date desc")
    try:
        response = requests.get(url, auth=DATA_AUTH)
        if response.status_code == 200:
            result = response.json()['value'][0]
    except:
        print(f"Возникла ошибка при получении данных из url: {url}")

    finally:
        return result



if __name__ == '__main__':
    get_ttn_details('8b108495-507d-11ee-8195-001dd8b72b55')
