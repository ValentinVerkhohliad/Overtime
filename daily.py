import glob
import os
import win32com.client
from lists import cs_dict

errors = ''


def remove_sales():
    k = 100
    while k > 2:
        val = sheet.Cells(k, 1)
        if str(val) == 'None' or str(val) == 'Total ':
            k -= 1
        else:
            val = sheet.Cells(k, 1).value.split(' ')[0]
            if val in cs_dict:
                sheet.Cells(k, 1).value = val
            else:
                sheet.Rows(k).EntireRow.Delete()
            k -= 1


def execute_():
    global sheet
    global errors
    xlsx_files = glob.glob1('C:\\reports\\dailyreports\\', '*.xlsx')
    xlApp = win32com.client.Dispatch("Excel.Application")
    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join('C:\\reports\\dailyreports\\', file))
        xlApp.Workbooks.Application.DisplayAlerts = False
        sheet = xlApp.ActiveSheet
        remove_sales()
        xlApp.Save()
        errors = errors + ('editing %s complete' % file + '\n')
    xlApp.Quit()
