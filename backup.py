import glob
import calendar
import shutil
import os


def make_dir():
    try:
        os.makedirs(logins_dest_path)
        os.makedirs(report_dest_path)
    except FileExistsError:
        pass


def copy_logins_files():
    try:
        shutil.copyfile(os.path.join(logins_source_path, file), os.path.join(logins_dest_path, file))
    except FileNotFoundError:
        pass


def copy_report_files():
    try:
        shutil.copyfile(os.path.join(report_source_path, 'Ot.xlsx'), os.path.join(report_dest_path, 'Ot.xlsx'))
        shutil.copyfile(os.path.join(report_source_path, 'Work_days.xlsx'),
                        os.path.join(report_dest_path, 'Work_days.xlsx'))
    except FileNotFoundError:
        pass


def clean_source_dir():
    try:
        os.remove(logins_source_path + '\\' + file)
    except FileNotFoundError:
        pass


def execute_():
    global file
    global report_dest_path
    global report_source_path
    global logins_source_path
    global logins_dest_path
    files = glob.glob1('C:\\reports\\logins\\', '*.xlsx')
    report_source_path = 'C:\\reports\\reports'
    logins_source_path = 'C:\\reports\\logins'
    try:
        logins_dest_path = 'C:\\reports\\backup\\' + files[0][18:28].split('.')[2] + '\\' \
                           + calendar.month_name[int(files[0][18:28].split('.')[1].lstrip('0'))] + '\\' + 'Logins'
        report_dest_path = 'C:\\reports\\backup\\' + files[0][18:28].split('.')[2] + '\\' \
                           + calendar.month_name[int(files[0][18:28].split('.')[1].lstrip('0'))] + '\\' + 'Reports'
    except IndexError:
        logins_dest_path = 'Error'
        print('Folder is empty or wrong file in it')
    make_dir()
    copy_report_files()
    if logins_dest_path != 'Error':
        for file in files:
            copy_logins_files()
            clean_source_dir()


