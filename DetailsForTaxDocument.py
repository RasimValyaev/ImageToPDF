import os
import sys
from datetime import datetime
import requests

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"D:\Prestige\Python\NovaPoshta"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"c:\Rasim\Python\NovaPoshta"
sys.path.append(os.path.abspath(CONFIG_PATH))

from configPrestige import DATA_AUTH


def get_counterparty(tax_code):
    counterparty = ''
    url = (r"http://192.168.1.254/utp_prestige/odata/standard.odata/Catalog_Контрагенты?$format=json&$"
           fr"filter=КодПоЕДРПОУ eq '{tax_code}'&$select=Ref_Key,Description")
    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        counterparty = resp.json()['value'][0].values()

    return counterparty


def get_doc_sale_details(uuid):
    result = ''
    url = (
        r"http://192.168.1.254/utp_prestige/odata/standard.odata/Document_РеализацияТоваровУслуг/?$format=json"
        r"&$filter=Posted eq true and ОтражатьВУправленческомУчете eq true "
        fr"and Ref_Key eq guid'{uuid}'&$select=Number,Date")
    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        result = resp.json()['value'][0]

    return result


def get_contract_details(uuid):
    result = ''
    url = (r"http://192.168.1.254/utp_prestige/odata/standard.odata/Catalog_ДоговорыКонтрагентов?$format=json"
           fr"&$inlinecount=allpages&$filter=Ref_Key eq guid'{uuid}'"
           r"&$select=Description,_НКС_ДнівВідтермінуванняОплати,Номер,Дата")
    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        result = resp.json()['value'][0]

    return result


def get_list_of_tax_fatura(date_sale, client_uuid):
    result = ''
    date_sale = datetime.strptime(date_sale, "%d.%m.%Y")
    day = date_sale.day
    month = date_sale.month
    year = date_sale.year

    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_НалоговаяНакладная?$format=json"
           f"&$orderby=Date desc&$filter=Контрагент_Key eq guid'{client_uuid}'"
           f"and year(Date) eq {year} "
           f"and month(Date) eq {month} "
           f"and day(Date) eq {day}"
           "&$select=Number,Date,ДокументОснование,ДоговорКонтрагента_Key")

    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        result = resp.json()['value']

    return result


if __name__ == '__main__':
    tax_code = '425477026551'
    get_counterparty(tax_code)
