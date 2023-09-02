# pip install pyxlsb
import pandas as pd
from datetime import datetime


def load_excel(path_to_excel):
    df = pd.read_excel(path_to_excel, sheet_name="Sheet1")
    return df


def create_index(df):
    vkl = 1
    i = 0
    last_date = datetime.strptime(df.loc[len(df) - 1, 'дата'], "%d.%m.%Y").date()
    if last_date < datetime.today().date():
        df.loc[len(df), 'дата'] = datetime.strftime(datetime.today().date(), "%d.%m.%Y")
        df.loc[len(df)-1, 'остаток'] = df.loc[len(df) - 2, 'остаток']

    for i, row in df.iterrows():
        df.loc[i, 'index'] = vkl
        if df.loc[i, 'остаток'] <= 2:
            if df.loc[i - 1, 'остаток'] <= 2:
                df.loc[i, 'index'] = 0
            else:
                vkl = vkl + 1

    print(df)


if __name__ == '__main__':
    excel_path = r"C:\Users\Rasim\Desktop\quantity_days.xlsb"
    df = load_excel(excel_path)
    create_index(df)
