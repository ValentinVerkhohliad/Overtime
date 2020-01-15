import win32com.client
import glob
import os

errors = ''
date_column = {'01': [2, 3, 4], '02': [5, 6, 7], '03': [8, 9, 10], '04': [11, 12, 13], '05': [14, 15, 16],
               '06': [17, 18, 19], '07': [20, 21, 22], '08': [23, 24, 25], '09': [26, 27, 28], '10': [29, 30, 31],
               '11': [32, 33, 34], '12': [35, 36, 37], '13': [38, 39, 40], '14': [41, 42, 43], '15': [44, 45, 46],
               '16': [47, 48, 49], '17': [50, 51, 52], '18': [53, 54, 55], '19': [56, 57, 58], '20': [59, 60, 61],
               '21': [62, 63, 64], '22': [65, 66, 67], '23': [68, 69, 70], '24': [71, 72, 73], '25': [74, 75, 76],
               '26': [77, 78, 79], '27': [80, 81, 82], '28': [83, 84, 85], '29': [86, 87, 88], '30': [89, 90, 91],
               '31': [92, 93, 94]}


def get_agent_row(agent):
    global errors
    try:
        for row in range(4, 50):
            if agent == daily.Sheets(1).Cells(row, 1).value:
                return row
    except Exception:
        errors = errors + 'excess agent in file or something went wrong \n'


def copy_total_calls():
    global errors
    i = 3
    agent = sheet.Cells(i, 1).value
    while str(agent) != 'Total ':
        call_type = sheet.Cells(i, 5).value
        call_count = sheet.Cells(i, 2).value
        if 'Incoming' in call_type:
            daily.Sheets(1).Cells(get_agent_row(agent), date_column[get_day()][0]).value = call_count
        elif 'Internal' in call_type:
            daily.Sheets(1).Cells(get_agent_row(agent), date_column[get_day()][2]).value = call_count
        elif 'Outgoing' in call_type:
            daily.Sheets(1).Cells(get_agent_row(agent), date_column[get_day()][1]).value = call_count
        else:
            errors = errors + 'This document is terrible, check it \n'
        i += 1
        agent = sheet.Cells(i, 1).value


def get_day():
    return file.lstrip('Masters_').split('.')[0]


def execute_():
    global sheet
    global file
    global daily
    global errors
    xlsx_files = glob.glob1('C:\\reports\\dailyreports\\', '*.xlsx')
    if len(xlsx_files) == 0:
        errors = errors + 'No XLSX files to convert. \n'
    xlApp = win32com.client.Dispatch("Excel.Application")
    daily = xlApp.Workbooks.Open(r'C:\reports\reports\Daily.xlsx')
    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join('C:\\reports\\dailyreports\\', file))
        xlApp.Workbooks.Application.DisplayAlerts = False
        sheet = xlApp.ActiveSheet
        copy_total_calls()
        xlApp.Save()
        errors = errors + 'editing %s complete' % file + '\n'
    xlApp.Quit()
