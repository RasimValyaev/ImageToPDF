# pip install pyxlsb
import pandas as pd
from datetime import datetime
import os
import sys

from authorize import sql_to_dataframe

if os.environ['COMPUTERNAME'] == 'PRESTIGEPRODUCT':
    CONFIG_PATH = r"d:\Prestige\Python\Config"
else:
    CONFIG_PATH = r"c:\Rasim\Python\Config"

sys.path.append(os.path.abspath(CONFIG_PATH))


def load_excel(path_to_excel):
    df = pd.read_excel(path_to_excel, sheet_name="Sheet1")
    return df


def create_index(df):
    vkl = 1
    df['index'] = None
    # last_date = datetime.strptime(df.loc[len(df) - 1, 'doc_date'], "%d.%m.%Y").date()
    last_date = df['doc_date'].iloc[-1].date()
    if last_date < datetime.today().date():
        df.loc[len(df), 'doc_date'] = datetime.strftime(datetime.today().date(), "%d.%m.%Y")
        df.loc[len(df) - 1, 'stok'] = df.loc[len(df) - 2, 'stok']

    for i, row in df.iterrows():
        df.loc[i, 'index'] = vkl
        if df.loc[i, 'stok'] <= 2:
            if df.loc[i - 1, 'stok'] <= 2:
                df.loc[i, 'index'] = 0
            else:
                vkl = vkl + 1

    print(df)


def get_df_from_table():
    sql = "SELECT * FROM public.t_one_stok ORDER BY doc_date desc Limit 100"
    df = sql_to_dataframe(sql)
    return df


if __name__ == '__main__':
    # excel_path = r"C:\Users\Rasim\Desktop\quantity_days.xlsb"
    # df = load_excel(excel_path)
    df = get_df_from_table()
    create_index(df)
