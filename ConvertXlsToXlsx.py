from pathlib import Path
import win32com.client as win32


def convert_xls_to_xlsx(xls_path) -> None:
    path = Path(xls_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(str(path.absolute().with_suffix(".xlsx")), FileFormat=51)
    wb.Close()
    excel.Application.Quit()
    return str(path.absolute().with_suffix(".xlsx"))

if __name__ == '__main__':
    xls_path = r"c:\Users\Rasim\Desktop\Scan\ТОВ ЛЕГІОН 2015\Написать письмо\Копия ЛЕГІОН 2015.xls"
    convert_xls_to_xlsx(xls_path)
