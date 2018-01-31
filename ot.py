#python3.5
import win32com.client
import time
import datetime
import glob
import os
import calendar

C = 3
D = 4
NINE_AM = datetime.datetime.strptime('9:00:00', '%H:%M:%S')
MOT = datetime.datetime.strptime('10:00:00', '%H:%M:%S')
SIX_PM = datetime.datetime.strptime('18:00:00', '%H:%M:%S')
SIX_THIRTY = datetime.datetime.strptime('18:30:00', '%H:%M:%S')
LOW_THRESHOLD = datetime.datetime.strptime('9:32:00', '%H:%M:%S')
WORK_DAY = SIX_PM - MOT
ONE_HOUR = MOT - NINE_AM
cs_list = ['70692', '70610', '70625', '70697', '70629', '70684', '70612', '70675', '70677', '70672', '70654', '70614',
           '70694', '70650', '70655', '70678', '70653', '70602', '70646', '70656', '70669', '70627', '70621', '70671',
           '70676', '70652', '70605', '70674']
morning_workers_list = ['70654', '70627']
half_day_list = ['70627', '70694', '70671', '70654']
sales_list = ['70683', '70682', '70616', '70611', '70607', '70680', '70689', '70673', '70670', '70679', '70662',
              '70601', '70644', '70604', '70651', '70664', '70622', '70665', '70695', '70659', '70686', '70626',
              '70609',
              '70645', '70681', '70623', '70696', '70615', '70613', '70638']
cs_dict = {'Abdel': '70692', 'Andre': '70625', 'Andreas': '70697', 'Bettina': '70684', 'Christine': '70612',
           'Dolores': '70675', 'Fabian': '70677', 'Guilia': '70672', 'Jana': '70654', 'Karel': '70694', 'Leoni': '70655',
           'Linda': '70653', 'Manuela': '70602', 'Marcel': '70646', 'Markus': '70656', 'Mathilde': '70669',
           'Mats': '70627', 'Mia': '70621', 'Nathalie': '70671', 'Oliver': '70676', 'Otto': '70652', 'Silvia': '70605',
           'Tony': '70674'}


def sort_and_format():
    xlAscending = 1
    xlSortColumns = 1
    xlApp.Sheets(1).Range("A2:D60").Sort(Key1=xlApp.Sheets(1).Range("B2"),
                                         Order1=xlAscending, Orientation=xlSortColumns)
    xlApp.Sheets(1).Columns('E:F').NumberFormat = "@"


def get_week_day(my_date):
    return calendar.day_name[my_date.weekday()]


def time_change(x):
    i = 2
    while i < 60:
        try:
            val = sheet.Cells(i, x).value
            val = str(val)
            val = val.split()
            sheet.Cells(i, x).NumberFormat = "@"
            sheet.Cells(i, x).value = val[1][0:8]
            i += 1
        except:
            break


def calculate_ot():
    global file
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = get_week_day(my_date)
    i = 2
    val = sheet.Cells(i, 1).value
    while val:
        val = sheet.Cells(i, 1).value
        if val in cs_list:
            ot = '00:00:00'
            eot = '00:00:00'
            arrive_time = datetime.datetime.strptime(sheet.Cells(i, 3).value, '%H:%M:%S')
            leave_time = datetime.datetime.strptime(sheet.Cells(i, 4).value, '%H:%M:%S')
            if day == 'Saturday':
                sum_ot = leave_time - arrive_time
                sheet.Cells(i, 5).value = str(sum_ot)
            else:
                if NINE_AM > arrive_time:
                    arrive_time = NINE_AM
                if arrive_time < MOT:
                    if arrive_time < LOW_THRESHOLD:
                        ot = MOT - arrive_time
                        if val in morning_workers_list:
                            ot = '00:00:00'
                if leave_time > SIX_THIRTY:
                    eot = leave_time - SIX_PM
                working_time = leave_time - arrive_time
                ot = str(ot)
                ot = ot.split(':')
                eot = str(eot)
                eot = eot.split(':')
                sum_ot = datetime.timedelta(hours=int(ot[0]), minutes=int(ot[1]), seconds=int(ot[2])) \
                         + datetime.timedelta(hours=int(eot[0]), minutes=int(eot[1]), seconds=int(eot[2]))
                if str(sum_ot) == '0:00:00':
                    pass
                else:
                    if (working_time - WORK_DAY) > ONE_HOUR and (working_time - WORK_DAY) > sum_ot:
                        sum_ot = working_time - WORK_DAY
                    sheet.Cells(i, 5).value = str(sum_ot)
        i += 1


def calculate_late():
    global file
    my_date = datetime.datetime.strptime(file[18:28], '%d.%m.%Y')
    day = get_week_day(my_date)
    i = 2
    luckers = []
    val = sheet.Cells(i, 1).value
    if day == 'Saturday' or day == 'Sunday':
        pass
    else:
        if day == 'Friday':
            lucky_people = input('Please input lucky people at %s. For example (Abdel,Andre,Andreas)' % file)
            lucky_people = lucky_people.split(',')
            try:
                for worker in lucky_people:
                    luckers.append(cs_dict[worker])
            except KeyError:
                print('Wrong name')
        while val:
            val = sheet.Cells(i, 1).value
            if val in cs_list:
                m_late = '0:00:00'
                e_late = '0:00:00'                
                arrive_time = datetime.datetime.strptime(sheet.Cells(i, 3).value, '%H:%M:%S')
                leave_time = datetime.datetime.strptime(sheet.Cells(i, 4).value, '%H:%M:%S')
                working_time = leave_time - arrive_time
                if arrive_time > MOT:
                    m_late = arrive_time - MOT
                    mia = m_late
                if val in luckers:
                    sum_late = m_late
                else:
                    if leave_time < SIX_PM:
                        e_late = SIX_PM - leave_time
                        if val in half_day_list:
                            e_late = '0:00:00'
                    m_late = str(m_late)
                    m_late = m_late.split(':')
                    e_late = str(e_late)
                    e_late = e_late.split(':')                          
                    sum_late = datetime.timedelta(hours=int(m_late[0]), minutes=int(m_late[1]), seconds=int(m_late[2]))\
                             + datetime.timedelta(hours=int(e_late[0]), minutes=int(e_late[1]), seconds=int(e_late[2]))
                    if day == 'Wednesday' and val == '70621':
                        sum_late = mia
                if str(sum_late) == '0:00:00' or sum_late > WORK_DAY or working_time >= WORK_DAY:
                    pass                
                else:
                    sheet.Cells(i, 6).value = str(sum_late)
            i += 1


def delete_sales():
    k = 60
    while k > 1:
        val = sheet.Cells(k, 1).value
        if val in sales_list:
            sheet.Rows(k).EntireRow.Delete()
        k -= 1


xlsx_files = glob.glob('*.xlsx')

if len(xlsx_files) == 0:
    raise RuntimeError('No XLSX files to convert.')
xlApp = win32com.client.Dispatch("Excel.Application")
for file in xlsx_files:
    xlWb = xlApp.Workbooks.Open(os.path.join(os.getcwd(), file))
    xlApp.Workbooks.Application.DisplayAlerts = False
    sheet = xlApp.ActiveSheet
    sort_and_format()
    time_change(C)
    time_change(D)
    calculate_ot()
    calculate_late()
    delete_sales()
    xlApp.Save()
    print('editing %s complete' % file)

xlApp.Quit()
time.sleep(2)
xlApp = None         