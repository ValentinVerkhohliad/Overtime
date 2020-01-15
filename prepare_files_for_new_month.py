import datetime
import calendar
import win32com.client

column_name = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I", 10: "J", 11: "K", 12: "L",
               13: "M", 14: "N", 15: "O", 16: "P", 17: "Q", 18: "R", 19: "S", 20: "T", 21: "U", 22: "V", 23: "W",
               24: "X", 25: "Y", 26: "Z", 27: "AA", 28: "AB", 29: "AC", 30: "AD", 31: "AE", 32: "AF", 33: "AG",
               34: "AH", 35: "AI", 36: "AJ", 37: "AK", 38: "AL", 39: "AM", 40: "AN", 41: "AO", 42: "AP", 43: "AQ",
               44: "AR", 45: "AS", 46: "AT", 47: "AU", 48: "AV", 49: "AW", 50: "AX", 51: "AY", 52: "AZ", 53: "BA",
               54: "BB", 55: "BC", 56: "BD", 57: "BE", 58: "BF", 59: "BG", 60: "BH", 61: "BI", 62: "BJ", 63: "BK",
               64: "BL", 65: "BM", 66: "BN", 67: "BO", 68: "BP", 69: "BQ", 70: "BR", 71: "BS", 72: "BT", 73: "BU",
               74: "BV", 75: "BW", 76: "BX", 77: "BY", 78: "BZ", 79: "CA", 80: "CB", 81: "CC", 82: "CD", 83: "CE",
               84: "CF", 85: "CG", 86: "CH", 87: "CI", 88: "CJ", 89: "CK", 90: "CL", 91: "CM", 92: "CN", 93: "CO",
               94: "CP", 95: "CQ", 96: "CR", 97: "CS", 98: "CT", 99: "CU", 100: "CV", 101: "CW", 102: "CX", 103: "CY",
               104: "CZ"}

errors = ''
weekend_formula = []
weekdays_formula = []


def color_of_weekends_and_ot_formulas(date, i, j, report_type):
    day = calendar.day_name[datetime.datetime.strptime(date, '%d.%m.%Y').weekday()]
    if day == 'Saturday' or day == 'Sunday':
        if report_type == daily:
            colour_range = column_name[i] + '4' + ':' + column_name[i + 2] + str(j - 1)
            report_type.Sheets(1).Range(colour_range).Interior.ColorIndex = 48
        else:
            colour_range = column_name[i] + '2' + ':' + column_name[i] + str(j - 1)
            report_type.Sheets(1).Range(colour_range).Interior.ColorIndex = 48
            if report_type == ot:
                weekend_formula.append(column_name[i] + '2')
    else:
        if report_type == ot:
            weekdays_formula.append(column_name[i] + '2')


def get_current_month():
    now = datetime.datetime.now()
    if now.month < 10:
        return '0' + str(now.month)
    else:
        return str(now.month)


def clear_range_(clear_range, report_type):
    report_type .Sheets(1).Range(clear_range).ClearContents()
    report_type.Sheets(1).Range(clear_range).Interior.TintAndShade = 0


def prepare_workdays_file():
    report_type = work_days
    j = 2
    """clear all previous month data"""
    val = work_days.Sheets(1).Cells(2, 1).value
    while val != None:
        j += 1
        val = work_days.Sheets(1).Cells(j, 1).value
    clear_range = "B2:AF" + str(j-1)
    clear_range_(clear_range, report_type)
    """change date in document to current month and make some colour"""
    for i in range(2, month_day_amount + 2):
        if i < 11:
            date = '0' + str(i - 1) + '.' + month + '.' + year
            work_days.Sheets(1).Cells(1, i).value = date
            color_of_weekends_and_ot_formulas(date, i, j, report_type)
        else:
            date = str(i - 1) + '.' + month + '.' + year
            work_days.Sheets(1).Cells(1, i).value = date
            color_of_weekends_and_ot_formulas(date, i, j, report_type)


def prepare_ot_file():
    report_type = ot
    j = 1
    """clear all previous month data"""
    val = ot.Sheets(1).Cells(2, 2).value
    while val != None:
        j += 1
        val = ot.Sheets(1).Cells(j, 2).value

    """clear ot range"""
    clear_range = "C2:AG" + str(j - 1)
    clear_range_(clear_range, report_type)

    """change date in document to current month and make some colour"""
    for i in range(3, month_day_amount + 3):
        if i < 12:
            date = '0' + str(i - 2) + '.' + month + '.' + year
            ot.Sheets(1).Cells(1, i).value = date
            color_of_weekends_and_ot_formulas(date, i, j, report_type)
        else:
            date = str(i - 2) + '.' + month + '.' + year
            ot.Sheets(1).Cells(1, i).value = date
            color_of_weekends_and_ot_formulas(date, i, j, report_type)

    """make formulas for next month"""
    ot_formula_range = 'AH2:AH' + str(j-1)
    wot_formula_range = 'AI2:AI' + str(j-1)
    ot.Sheets(1).Cells(2, 34).Formula = '=Sum(' + ','.join(weekdays_formula) + ')'
    ot.Sheets(1).Cells(2, 35).Formula = '=Sum(' + ','.join(weekend_formula) + ')'
    ot.Sheets(1).Range(ot_formula_range).Formula = ot.Sheets(1).Cells(2, 34).Formula
    ot.Sheets(1).Range(wot_formula_range).Formula = ot.Sheets(1).Cells(2, 35).Formula

    """copy date for late section"""
    while str(val) != 'Date':
        j += 1
        val = ot.Sheets(1).Cells(j, 2).value
    paste_range = 'C' + str(j) + ':' + 'AG' + str(j)
    ot.Sheets(1).Range(paste_range).value = ot.Sheets(1).Range("C1:AG1").value
    y = j + 1
    while val!= None:
        j += 1
        val = ot.Sheets(1).Cells(j, 2).value
    """clear all working range from previous data"""
    clear_range = "C" + str(y) + ':' + 'AG' + str(j - 1)
    clear_range_(clear_range, report_type)


def prepare_daily_file():
    report_type = daily
    j = 4
    """clear all previous month data"""
    val = daily.Sheets(1).Cells(4, 1).value
    while val != None:
        j += 1
        val = daily.Sheets(1).Cells(j, 1).value
    clear_range = "B4:CP" + str(j - 1)
    clear_range_(clear_range, report_type)
    """change date in document to current month and make some colour"""
    k = 1
    for i in range(2, month_day_amount * 3 + 2, 3):
        if i < 28:
            date = '0' + str(k) + '.' + month + '.' + year
            daily.Sheets(1).Cells(1, i).value = date
            k += 1
            color_of_weekends_and_ot_formulas(date, i, j, report_type)
        else:
            date = str(k) + '.' + month + '.' + year
            daily.Sheets(1).Cells(1, i).value = date
            k += 1
            color_of_weekends_and_ot_formulas(date, i, j, report_type)


def execute_():
    global work_days
    global month
    global year
    global month_day_amount
    global ot
    global daily
    global errors
    month = get_current_month()
    year = str(datetime.datetime.now().year)
    month_day_amount = calendar.mdays[datetime.date.today().month]
    xlapp = win32com.client.Dispatch("Excel.Application")
    ot = xlapp.Workbooks.Open('C:\\reports\\reports\\Ot.xlsx')
    work_days = xlapp.Workbooks.Open(r'C:\reports\reports\Work_days.xlsx')
    daily = xlapp.Workbooks.Open('C:\\reports\\reports\\Daily.xlsx')
    xlapp.Workbooks.Application.DisplayAlerts = False
    prepare_workdays_file()
    prepare_ot_file()
    prepare_daily_file()
    xlapp.Save()
    xlapp.Quit()
