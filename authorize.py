import os
import sys
import traceback
import pandas as pd
from sqlalchemy import create_engine
import sqlalchemy as sa

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"D:\Prestige\Python\NovaPoshta"
else:
    CONFIG_PATH = r"C:\Rasim\Python\Config"
    CONFIG_PATH_NOVAPOSHTA = r"c:\Rasim\Python\NovaPoshta"

sys.path.append(os.path.abspath(CONFIG_PATH))
from configPrestige import username, psw, hostname, port, basename, URL_CONST, chatid_rasim, DATA_AUTH, schema

engine = create_engine('postgresql://%s:%s@%s:%s/%s' % (username, psw, hostname, port, basename), pool_pre_ping=True,
                       connect_args={
                           "keepalives": 1,
                           "keepalives_idle": 30,
                           "keepalives_interval": 10,
                           "keepalives_count": 5,
                       })

def sql_to_dataframe(sql_query, odata=''):
    # conname - тип соединения. Результат запроса переводит в pandas DataFrame
    sms = ''
    df = ''
    try:
        with engine.begin() as conn:
            if odata == '':
                df = pd.read_sql(sa.text(sql_query), conn)
            else:
                df = pd.read_sql_query(sa.text(sql_query), conn, params=odata)

    except Exception as e:
        sms = traceback.print_exc()
        print(str(sms))
        sms = "ERROR:ConnectToBase:dfExtract: %s" % sms


    finally:
        return df
