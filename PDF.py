import win32com.client as win32
import win32com
from PIL import ImageGrab
from PIL import Image
import win32api
import pythoncom
from functools import lru_cache
import pathlib

def PDF(Get):
    # Path to original excel file
    excelfile = (r"PAYROLL.xlsx")
    # PDF path when saving
    pdf_file = (r"PAYROLL" + Get +".pdf")
    excel_path = str(pathlib.Path.cwd() / excelfile)
    pdf_path = str(pathlib.Path.cwd() / pdf_file)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = 0
    excel.DisplayAlerts = 0

    try:
        print('Start conversion to PDF')
        # Open
        wb = excel.Workbooks.Open(excel_path)
        ws_index_list = [2]
        # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
        wb.WorkSheets(ws_index_list).Select()


        # Save
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

    except Exception as e:
        print('failed.'% e)
        print(str(e))
    else:
        wb.SaveAs(pdf_path, FileFormat=57)
        print('Succeeded.')
    finally:
        wb.Close(False)
        excel.Quit()

        wb = None
        excel = None
    print('finish')
    

