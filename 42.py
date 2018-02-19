import win32com.client
import time
import glob
import os
import datetime
import calendar


def get_week_day(my_date):
    return calendar.day_name[my_date.weekday()]


def adjust_time(x):
    """adjusting time"""
    try:
        hours, minutes, sec = x.split(':')
        sec = ':00'
        if int(hours) == 0:
            if 25 <= int(minutes) <= 50:
                minutes = ':30'
            elif int(minutes) < 25:
                minutes = ':00'
            elif int(minutes) > 50:
                minutes = ':00'
                h = int(hours) + 1
                hours = str(h)
        else:
            hours += ':'
        return hours + minutes + sec
    except:
        pass


def ot():
    """create workers dict and their ot"""
    global sheet
    sot = {}
    for i in range(2, 26):
        val = sheet.Cells(i, 1).value
        tot = sheet.Cells(i, 5).value
        try:
            tot = adjust_time(tot)
        except AttributeError:
            pass
        sot[val] = tot
    return sot


def late():
    """create workers dict and their late time"""
    global sheet
    dic = {}
    for i in range(2, 26):
        val = sheet.Cells(i, 1).value
        lat = sheet.Cells(i, 6).value
        dic[val] = lat
    return dic


def get_col():
    """create column dictionary"""
    col_dict = {}
    for i in range(3, 35):
        column_ = report.Sheets(1).Cells(1, i).value
        column_ = str(column_).split()
        column_ = column_[0].split('-')
        column_ = str(column_[2])
        col_dict[column_] = i
    return col_dict


def copy_ot():
    global report
    worker_list = ot()
    date = file.split('.')
    date = str(date[0].lstrip('MastersDailyLogins'))
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = get_week_day(my_date)
    if day == 'Sunday':
        pass
    else:
        if date in column:
            col = column[date]
            for i in range(2, 26):
                try:
                    row = report.Sheets(1).Cells(i, 1).value
                    report.Sheets(1).Cells(i, col).value = worker_list[row]
                except KeyError:
                    pass
        print('copy ot in %s complete' % file)
        report.Save()


def copy_late():
    global report
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = get_week_day(my_date)
    if day == 'Saturday' or day == 'Sunday':
        pass
    else:
        worker_list = late()
        date = file.split('.')
        date = str(date[0].lstrip('MastersDailyLogins'))
        if date in column:
            col = column[date]
            for i in range(32, 56):
                try:
                    row = report.Sheets(1).Cells(i, 1).value
                    report.Sheets(1).Cells(i, col).value = worker_list[row]
                except KeyError:
                    pass
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
