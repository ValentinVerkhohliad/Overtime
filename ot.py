#python3.5
import time
import datetime
import glob
import os
import calendar
import win32com.client
from lists import cs_list, half_day_list, sales_list, cs_dict

errors = ''
NINE_AM = datetime.datetime.strptime('9:00:00', '%H:%M:%S')
TWO_FIFTEEN = datetime.datetime.strptime('14:15:00', '%H:%M:%S')
FALSE_HOUR = datetime.datetime.strptime('9:08:00', '%H:%M:%S')
MOT = datetime.datetime.strptime('10:00:00', '%H:%M:%S')
SIX_PM = datetime.datetime.strptime('18:00:00', '%H:%M:%S')
SIX_THIRTY = datetime.datetime.strptime('18:28:00', '%H:%M:%S')
LOW_THRESHOLD = datetime.datetime.strptime('9:30:00', '%H:%M:%S')
TWELVE = datetime.datetime.strptime('12:00:00', '%H:%M:%S')
OT_DAY = SIX_THIRTY - MOT
WORK_DAY = SIX_PM - MOT
HALF_DAY = TWO_FIFTEEN - MOT
TWO_HOUR = TWELVE - MOT
ONE_HOUR = MOT - FALSE_HOUR
HALF_HOUR = MOT - LOW_THRESHOLD
friday_count = 0


def sort_and_format():
    xlascending = 1
    xlsortcolumns = 1
    xlApp.Sheets(1).Range("A2:D60").Sort(Key1=xlApp.Sheets(1).Range("B2"),
                                         Order1=xlascending, Orientation=xlsortcolumns)
    xlApp.Sheets(1).Columns('E:F').NumberFormat = "@"


def time_change():
    for k in range(3, 5):
        for i in range(2, 60):
            try:
                val = sheet.Cells(i, k).value
                val = str(val)
                val = val.split()
                sheet.Cells(i, k).NumberFormat = "@"
                sheet.Cells(i, k).value = val[1][0:8]
            except IndexError:
                break


def get_day():
    global day
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = calendar.day_name[my_date.weekday()]


def luckers_list(lucky_people):
    global luckers
    global friday_count
    global errors
    luckers = []
    if day == 'Friday':
        try:
            for worker in lucky_people[friday_count]:
                luckers.append(cs_dict[worker])
                friday_count += 1
        except KeyError:
            errors = errors + 'Wrong name in luckers or field is empty\n'
            friday_count += 1
    return luckers


def calculate_ot():
    global file
    i = 2
    val = sheet.Cells(i, 1).value
    while val != None:
        val = sheet.Cells(i, 1).value
        if val in cs_list:
            arrive_time = datetime.datetime.strptime(sheet.Cells(i, 3).value, '%H:%M:%S')
            leave_time = datetime.datetime.strptime(sheet.Cells(i, 4).value, '%H:%M:%S')
            if day == 'Saturday':
                sum_ot = leave_time - arrive_time
                sheet.Cells(i, 5).value = str(sum_ot)
            else:
                if NINE_AM > arrive_time:
                    arrive_time = NINE_AM
                if val in luckers or leave_time == '00:00:00':
                    leave_time = SIX_PM
                working_time = leave_time - arrive_time
                if val in half_day_list:
                    if (working_time - HALF_DAY) > HALF_HOUR:
                        sum_ot = working_time - HALF_DAY
                        sheet.Cells(i, 5).value = str(sum_ot)
                else:
                    if working_time < OT_DAY:
                        pass
                    else:
                        if (working_time - WORK_DAY) > ONE_HOUR:
                            sum_ot = working_time - WORK_DAY
                        else:
                            if (MOT - arrive_time) > HALF_HOUR or (leave_time - SIX_PM) > HALF_HOUR:
                                sum_ot = HALF_HOUR
                        sheet.Cells(i, 5).value = str(sum_ot)
        i += 1


def calculate_late():
    global file
    if day == 'Saturday':
        pass
    else:
        i = 2
        val = sheet.Cells(i, 1).value
        while val != None:
            val = sheet.Cells(i, 1).value
            if val in cs_list:
                if val == '70692':
                    pass
                else:
                    arrive_time = datetime.datetime.strptime(sheet.Cells(i, 3).value, '%H:%M:%S')
                    leave_time = datetime.datetime.strptime(sheet.Cells(i, 4).value, '%H:%M:%S')
                    if val in luckers:
                        leave_time = SIX_PM
                    working_time = leave_time - arrive_time
                    if val in half_day_list:
                        if (day == 'Thursday' or day == 'Friday') and val == '70694':
                            if working_time < TWO_HOUR:
                                sum_late = TWO_HOUR - working_time
                                sheet.Cells(i, 6).value = str(sum_late)
                        else:
                            if working_time < HALF_DAY:
                                sum_late = HALF_DAY - working_time
                                sheet.Cells(i, 6).value = str(sum_late)
                    else:
                        if working_time < WORK_DAY:
                            sum_late = WORK_DAY - working_time
                            sheet.Cells(i, 6).value = str(sum_late)
            i += 1


def delete_sales():
    k = 60
    while k > 1:
        val = sheet.Cells(k, 1).value
        if val in sales_list:
            sheet.Rows(k).EntireRow.Delete()
        k -= 1


def execute_(lucky_people):
    global xlWb
    global sheet
    global xlApp
    global file
    global errors
    xlsx_files = glob.glob1('C:\\reports\\logins\\', '*.xlsx')
    xlApp = win32com.client.Dispatch("Excel.Application")
    for file in xlsx_files:
        xlWb = xlApp.Workbooks.Open(os.path.join('C:\\reports\\logins\\', file))
        xlApp.Workbooks.Application.DisplayAlerts = False
        sheet = xlApp.ActiveSheet
        sort_and_format()
        time_change()
        get_day()
        if day == 'Sunday':
            pass
        else:
            luckers_list(lucky_people)
            calculate_ot()
            calculate_late()
            delete_sales()
        xlApp.Save()
        errors = errors + ('editing %s complete' % file + '\n')
    xlApp.Quit()
    time.sleep(2)
    xlApp = None
