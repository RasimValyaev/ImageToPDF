import os
from pathlib import Path
import win32com.client as win32


def convert_xls_to_xlsx(xls_path) -> str:
    path = Path(xls_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    save_to_xlsx_path = str(path.absolute().with_suffix(".xlsx"))
    if os.path.isfile(save_to_xlsx_path):
        os.remove(save_to_xlsx_path)
    wb.SaveAs(save_to_xlsx_path, FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    return save_to_xlsx_path


if __name__ == '__main__':
    xls_path = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЛЕГІОН 2015\Написать письмо\Копия ЛЕГІОН 2015.xls"
    convert_xls_to_xlsx(xls_path)
