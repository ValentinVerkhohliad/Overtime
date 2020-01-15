import glob
import calendar
import shutil
import os
import win32com.client

backup_error = ''
report_source_path = 'C:\\reports\\reports'
logins_source_path = 'C:\\reports\\logins'
daily_source_path = 'C:\\reports\\dailyreports'


def make_dir():
    try:
        os.makedirs(daily_dest_path)
        os.makedirs(logins_dest_path)
        os.makedirs(report_dest_path)
    except FileExistsError:
        pass


def copy_logins_files():
    try:
        shutil.copyfile(os.path.join(logins_source_path, file), os.path.join(logins_dest_path, file))
    except FileNotFoundError:
        pass


def copy_daily_files():
    try:
        shutil.copyfile(os.path.join(daily_source_path, file), os.path.join(daily_dest_path, file))
    except FileNotFoundError:
        pass


def copy_report_files():
    try:
        shutil.copyfile(os.path.join(report_source_path, 'Ot.xlsx'), os.path.join(report_dest_path, 'Ot.xlsx'))
        shutil.copyfile(os.path.join(report_source_path, 'Work_days.xlsx'),
                        os.path.join(report_dest_path, 'Work_days.xlsx'))
        shutil.copyfile(os.path.join(report_source_path, 'Daily.xlsx'), os.path.join(report_dest_path, 'Daily.xlsx'))
        shutil.copyfile(os.path.join(report_source_path, glob.glob1(report_source_path, '*.xlsx')[-1]),
                        os.path.join(report_dest_path, glob.glob1(report_source_path, '*.xlsx')[-1]))
    except FileNotFoundError:
        pass


def clean_logins_source_dir():
    try:
        os.remove(logins_source_path + '\\' + file)
    except FileNotFoundError:
        pass


def clean_daily_source_dir():
    try:
        os.remove(daily_source_path + '\\' + file)
    except FileNotFoundError:
        pass


def execute_():
    global file
    global report_dest_path
    global logins_dest_path
    global backup_error
    global daily_dest_path
    xlApp = win32com.client.Dispatch("Excel.Application")
    report = xlApp.Workbooks.Open(r'C:\reports\reports\Ot.xlsx')
    date = str(report.Sheets(1).Cells(1, 3).value).lstrip(' ')
    year = date[0:4]
    month = date[5:7]
    xlApp.Quit()
    files = glob.glob1('C:\\reports\\logins\\', '*.xlsx')
    daily_files = glob.glob1('C:\\reports\\dailyreports\\', '*.xlsx')

    try:
        logins_dest_path = 'C:\\reports\\backup\\' + year + '\\' + calendar.month_name[int(month)] + '\\' + 'Logins'
        report_dest_path = 'C:\\reports\\backup\\' + year + '\\' + calendar.month_name[int(month)] + '\\' + 'Reports'
        daily_dest_path = 'C:\\reports\\backup\\' + year + '\\' + calendar.month_name[int(month)] + '\\' + 'Daily'
    except Exception:
        logins_dest_path = 'Error'
        backup_error = 'Something went wrong \n'

    if logins_dest_path != 'Error':
        make_dir()
        copy_report_files()
        for file in files:
            copy_logins_files()
            clean_logins_source_dir()
        for file in daily_files:
            copy_daily_files()
            clean_daily_source_dir()
