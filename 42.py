import win32com.client
import time
import glob
import os
import datetime
import calendar

def get_week_day(my_date):
    return calendar.day_name[my_date.weekday()]
    
    
def adjustTime(x):
    try:
        hours, min, sec = x.split(':')
        sec = ':00'
        if 25 <= int(min) <= 50:
            min = ':30'
        elif int(min) < 25:
            min = ':00'
        elif int(min) > 50:
            min = ':00'
            h = int(hours) + 1
            hours = str(h)
        return hours + min + sec
    except:
        print('Wrong time format for adjusting')

def ot():
    """create workers dict and their ot"""
    global sheet
    sot = {}
    i = 2
    val = sheet.Cells(i, 1).value
    while i < 25:
        val = sheet.Cells(i, 1).value
        tot = sheet.Cells(i, 5).value
        try:
            tot = adjustTime(tot)
        except AttributeError:
            pass
        sot[val] = tot
        i += 1
    return sot


def late():
    """create workers dict and their late time"""
    global sheet
    dic = {}
    i = 2
    val = sheet.Cells(i, 1).value
    while i < 25:
        val = sheet.Cells(i, 1).value
        lat = sheet.Cells(i, 6).value
        dic[val] = lat
        i += 1
    return dic


def get_col():
    col_dict = {}
    i = 3
    while i < 34:
        column = report.Sheets(1).Cells(1, i).value
        column = str(column).split()
        column = column[0].split('-')
        column = str(column[2])
        col_dict[column] = i
        i += 1
    return col_dict


def copy_ot():
    global report
    global file
    global column
    i = 2
    worker_list = ot()
    date = file.split('.')
    date = str(date[0].lstrip('MastersDailyLogins'))
    if date in column:
        col = column[date]
        while i < 25:
            try:
                row = report.Sheets(1).Cells(i, 1).value
                report.Sheets(1).Cells(i, col).value = worker_list[row]
                i += 1
            except KeyError:
                print('Somebody is absent today')
                i += 1
    print('copy ot in %s complete' % file)
    report.Save()


def copy_late():
    global report
    global file
    global column
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = get_week_day(my_date)
    if day == 'Saturday' or day == 'Sunday':
        pass    
    else:
        i = 32
        worker_list = late()
        date = file.split('.')
        date = str(date[0].lstrip('MastersDailyLogins'))
        if date in column:
            col = column[date]
            while i < 55:
                try:
                    row = report.Sheets(1).Cells(i, 1).value
                    report.Sheets(1).Cells(i, col).value = worker_list[row]
                    i += 1
                except KeyError:
                    i += 1
        print('copy late in %s complete' % file)
        report.Save()


xlsx_files = glob.glob('*.xlsx')
if len(xlsx_files) == 0:
    raise RuntimeError('No XLSX files to convert.')
xlApp = win32com.client.Dispatch("Excel.Application")
report = xlApp.Workbooks.Open('C:\\reports\\reports\\Ot.xlsx')
column = get_col()

for file in xlsx_files:
    xlWb = xlApp.Workbooks.Open(os.path.join(os.getcwd(), file))
    xlApp.Workbooks.Application.DisplayAlerts = False
    sheet = xlApp.ActiveSheet
    copy_ot()
    copy_late()
    xlApp.Save()


xlApp.Quit()
time.sleep(2)
xlApp = None