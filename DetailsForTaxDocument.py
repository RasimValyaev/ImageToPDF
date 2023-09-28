import os
import sys
import requests
from dateutil.parser import parse
from datetime import datetime

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"D:\Prestige\Python\NovaPoshta"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"c:\Rasim\Python\NovaPoshta"
sys.path.append(os.path.abspath(CONFIG_PATH))

from configPrestige import DATA_AUTH


def get_data_from_taxdoc(taxdoc_date, taxdoc_number, counterparty_taxnumber):
    counterparty = []
    year = parse(taxdoc_date, dayfirst=True).year
    month = parse(taxdoc_date, dayfirst=True).month
    date = parse(taxdoc_date, dayfirst=True).day
    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_НалоговаяНакладная?"
           f"$format=json&$filter=Posted eq true and endswith(Number, '{taxdoc_number}') eq true "
           f"and year(Date) eq {year} and month(Date) eq {month} and day(Date) eq {date} "
           f"and Контрагент/ИНН eq '{counterparty_taxnumber}'&$inlinecount=allpages&$expand=*")

    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        if len(resp.json()['value']) != 0:
            counterparty = resp.json()['value'][0]
        else:
            print("Не нашел клиента с ИНН ", counterparty_taxnumber)

    return counterparty


def get_doc_transport(doc_uuid, doc_type_is_sale='Document_РеализацияТоваровУслуг'):
    result = ''
    url = (r"http://192.168.1.254/utp_prestige/odata/standard.odata/Document_скТоварноТранспортнаяНакладная?"
           r"$expand=ДокументОснование&$format=json&$select=Date,Number,Ref_Key"
           f"&$filter=cast(guid'{doc_uuid}','{doc_type_is_sale}') eq ДокументОснование")
    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        if len(resp.json()['value']) != 0:
            result = resp.json()['value'][0]

    return result


def get_doc_sale_details(doc_uuid, doc_type_is_sale='Document_РеализацияТоваровУслуг'):
    result = ''
    url = ("http://192.168.1.254/utp_prestige/odata/standard.odata/Document_РеализацияТоваровУслуг/?"
           f"$format=json&$select=Date,Number,Сделка,Сделка_Type"
           f"&$inlinecount=allpages&$filter=Posted eq true and Ref_Key eq guid'{doc_uuid}'&$expand=*")
    resp = requests.get(url, auth=DATA_AUTH)
    if resp.status_code == 200:
        result = resp.json()['value'][0]

    return result


if __name__ == '__main__':
    tax_code = '425477026551'
    counterparty = get_counterparty_by_taxcode(tax_code)
    print(counterparty)
