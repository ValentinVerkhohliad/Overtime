import win32com.client
import time
import glob
import os
import datetime
import calendar
from lists import cs_dict


def get_day():
    global day
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = calendar.day_name[my_date.weekday()]


def get_date():
    return int(file.lstrip('MastersDailyLogins').split('.')[0])


def get_workers_set():
    worker_set = set(cs_dict)
    return worker_set


def get_today_workers_set():
    i = 2
    today_workers = []
    val = sheet.Cells(i, 2).value
    while val != None:
        try:
            val = sheet.Cells(i, 2).value
            val = val.split()[0]
            today_workers.append(val)
        except AttributeError:
            pass
        i += 1
    return set(today_workers)


def get_apsend_people_list():
    all_workers = get_workers_set()
    today_workers = get_today_workers_set()
    apsend_workers = all_workers.difference(today_workers)
    filtered_apsend_workers = filter(None, apsend_workers)
    apsend_workers = set(filtered_apsend_workers)
    return apsend_workers


def fill_visit_in_file():
    i = 2
    val = work_days.Sheets(1).Cells(i, 1).value
    while val != 'Billy':
        val = work_days.Sheets(1).Cells(i, 1).value
        if (day == 'Thursday' or day == 'Friday') and val == 'Karel':
            work_days.Sheets(1).Cells(i, get_date() + 1).value = 0.25
        else:
            if val == 'Karel' or val == 'Mats' or val == 'Zoe':
                work_days.Sheets(1).Cells(i, get_date() + 1).value = 0.5
            else:
                work_days.Sheets(1).Cells(i, get_date() + 1).value = 1
            if val in apsend_people_list:
                work_days.Sheets(1).Cells(i, get_date() + 1).value = 0
        i += 1


def execute_():
    global work_days
    global file
    global sheet
    global xlWb
    global xlApp
    global apsend_people_list
    xlsx_files = glob.glob1('C:\\reports\\logins\\', '*.xlsx')
    if len(xlsx_files) == 0:
        raise RuntimeError('No XLSX files to convert.')
    xlApp = win32com.client.Dispatch("Excel.Application")
    work_days = xlApp.Workbooks.Open(r'C:\reports\reports\Work_days.xlsx')
    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join('C:\\reports\\logins\\', file))
        xlApp.Workbooks.Application.DisplayAlerts = False
        sheet = xlApp.ActiveSheet
        get_day()
        if day == 'Saturday' or day == 'Sunday':
            pass
        else:
            apsend_people_list = get_apsend_people_list()
            fill_visit_in_file()
        print('editing %s complete' % file)
        xlApp.Save()
    xlApp.Quit()
    time.sleep(2)
    xlApp = None
