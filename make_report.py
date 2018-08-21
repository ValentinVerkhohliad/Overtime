import win32com.client
import time
import glob
import os
import datetime
import calendar

errors = ''


def get_day():
    global day
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = calendar.day_name[my_date.weekday()]


def get_late_start_row():
    global errors
    for i in range(20, 40):
        if report.Sheets(1).Cells(i, 1).value == '70697':
            return i
        elif i == 40:
            errors = errors + 'Please change start point for late row'


def ot():
    """create workers dict and their ot"""
    global sheet
    sot = {}
    i = 2
    val = sheet.Cells(i, 1).value
    while val != None:
        val = sheet.Cells(i, 1).value
        tot = sheet.Cells(i, 5).value
        if tot == None:
            pass
        else:
            sot[val] = tot
        i += 1
    return sot


def late():
    """create workers dict and their late time"""
    global sheet
    dic = {}
    i = 2
    val = sheet.Cells(i, 1).value
    while val != None:
        val = sheet.Cells(i, 1).value
        lat = sheet.Cells(i, 6).value
        if lat == None:
            pass
        else:
            dic[val] = lat
        i += 1
    return dic


def get_date():
    return int(file.lstrip('MastersDailyLogins').split('.')[0])


def copy_ot():
    global report
    worker_list = ot()
    if day == 'Sunday':
        pass
    else:
        i = 2
        row = report.Sheets(1).Cells(i, 1).value
        while row != None:
            try:
                row = report.Sheets(1).Cells(i, 1).value
                report.Sheets(1).Cells(i, get_date()+2).value = worker_list[row]
            except KeyError:
                pass
            i += 1


def copy_late():
    global report
    if day == 'Saturday' or day == 'Sunday':
        pass
    else:
        worker_list = late()
        i = late_start_point
        row = report.Sheets(1).Cells(i, 1).value
        while row != None:
            try:
                row = report.Sheets(1).Cells(i, 1).value
                report.Sheets(1).Cells(i, get_date() + 2).value = worker_list[row]
            except KeyError:
                pass
            i += 1


def execute_():
    global report
    global file
    global sheet
    global xlWb
    global xlApp
    global late_start_point
    global errors
    xlsx_files = glob.glob1('C:\\reports\\logins\\', '*.xlsx')
    if len(xlsx_files) == 0:
        raise RuntimeError('No XLSX files to convert.')
    xlApp = win32com.client.Dispatch("Excel.Application")
    report = xlApp.Workbooks.Open(r'C:\reports\reports\Ot.xlsx')
    late_start_point = get_late_start_row()
    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join('C:\\reports\\logins\\', file))
        xlApp.Workbooks.Application.DisplayAlerts = False
        sheet = xlApp.ActiveSheet
        get_day()
        copy_ot()
        copy_late()
        xlApp.Save()
        errors = errors + ('editing %s complete' % file + '\n')
    xlApp.Quit()
    time.sleep(2)
    xlApp = None
