import sys

from design import *
import gmail
import ot
import make_report
import work_days
import backup
import lists
import daily
import Excel_exit
import prepare_files_for_new_month
import make_daily_report
import collections
from PyQt5 import QtWidgets

month_dict = {'January': '.01', 'February': '.02', 'March': '.03', 'April': '.04', 'May': '.05', 'June': '.06',
              'July': '.07', 'August': '.08', 'September': '.09', 'October': '.10', 'November': '.11',
              'December': '.12'}


class MyWin(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.month = '01'
        self.label_type = 'py-cw1'
        self.prefix = "C:\\reports\\logins\\"
        self.ui.pushButton.clicked.connect(self.gmail_download)
        self.ui.comboBox.addItems(['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
                                   'September', 'October', 'November', 'December'])
        self.ui.comboBox_1.addItems(['Ot Logins', 'Daily'])
        self.ui.comboBox.activated[str].connect(self.month_selection)
        self.ui.comboBox_1.activated[str].connect(self.report_type_selection)
        self.ui.pushButton_2.clicked.connect(self.calculate_ot)
        self.ui.pushButton_3.clicked.connect(self.make_report_exec)
        self.ui.pushButton_4.clicked.connect(self.make_attendance)
        self.ui.pushButton_5.clicked.connect(self.make_backup)
        self.ui.pushButton_6.clicked.connect(self.add_to_CCA_cs_list)
        self.ui.pushButton_7.clicked.connect(self.delete_from_CCA_cs_list)
        self.ui.pushButton_8.clicked.connect(self.print_cs_cca_list)
        self.ui.pushButton_9.clicked.connect(self.add_to_half_day_CCA_cs_list)
        self.ui.pushButton_10.clicked.connect(self.delete_from_half_day_CCA_cs_list)
        self.ui.pushButton_13.clicked.connect(self.print_half_day_cs_cca_list)
        self.ui.pushButton_11.clicked.connect(self.add_to_sales_CCA_cs_list)
        self.ui.pushButton_12.clicked.connect(self.delete_from_sales_CCA_cs_list)
        self.ui.pushButton_14.clicked.connect(self.print_sales_cs_cca_list)
        self.ui.pushButton_15.clicked.connect(self.update_cs_dict_)
        self.ui.pushButton_16.clicked.connect(self.add_to_cs_dict)
        self.ui.pushButton_17.clicked.connect(self.delete_from_cs_dict)
        self.ui.pushButton_18.clicked.connect(self.print_cs_dict)
        self.ui.pushButton_19.clicked.connect(self.save_lists)
        self.ui.pushButton_20.clicked.connect(self.prepare_daily_files)
        self.ui.pushButton_21.clicked.connect(self.make_daily_report_function)
        self.ui.pushButton_22.clicked.connect(self.prepare_reports)
        self.ui.pushButton_23.clicked.connect(self.kill_excel)

    def gmail_download(self):
        gmail.execute_(self.month, self.label_type, self.prefix)
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(gmail.errors)
        self.ui.textBrowser.insertPlainText('Download complete \n')
        gmail.errors = ''

    def month_selection(self, text):
        self.month = month_dict[text]

    def report_type_selection(self, text):
        if text == 'Ot Logins':
            self.label_type = 'py-cw1'
            self.prefix = "C:\\reports\\logins\\"
        elif text == 'Daily':
            self.label_type = 'py-cw2'
            self.prefix = "C:\\reports\\dailyreports\\"

    def prepare_daily_files(self):
        daily.execute_()
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(daily.errors)
        self.ui.textBrowser.insertPlainText('Daily files prepared \n')

    def make_daily_report_function(self):
        make_daily_report.execute_()
        self.ui.textBrowser.clear()
        if make_daily_report.error != '':
            self.ui.textBrowser.insertPlainText(make_daily_report.error)
        else:
            self.ui.textBrowser.insertPlainText('Daily report is ready \n')

    def calculate_ot(self):
        lucky_people = []
        lucky_people.append(self.ui.lineEdit.text().split(','))
        lucky_people.append(self.ui.lineEdit_2.text().split(','))
        lucky_people.append(self.ui.lineEdit_3.text().split(','))
        lucky_people.append(self.ui.lineEdit_4.text().split(','))
        lucky_people.append(self.ui.lineEdit_14.text().split(','))
        ot.execute_(lucky_people)
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(ot.errors)
        self.ui.textBrowser.insertPlainText('Calculation Ot complete \n')

    def make_report_exec(self):
        make_report.execute_()
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(make_report.errors)
        self.ui.textBrowser.insertPlainText('Main Ot reropt is Done \n')
        make_report.errors = ''

    def make_attendance(self):
        work_days.execute_()
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(work_days.errors)
        self.ui.textBrowser.insertPlainText('Attendance reropt is Done \n')
        work_days.errors = ''

    def prepare_reports(self):
        prepare_files_for_new_month.execute_()
        if prepare_files_for_new_month.errors == '':
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText('All reports are ready for use. Enjoy\n')
            self.ui.textBrowser.insertPlainText(prepare_files_for_new_month.errors)
        else:
            self.ui.textBrowser.insertPlainText(prepare_files_for_new_month.errors)

    def make_backup(self):
        backup.execute_()
        if backup.backup_error == '':
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText('Backup is Done \n')
            self.ui.textBrowser.insertPlainText(backup.backup_error)
        else:
            self.ui.textBrowser.insertPlainText(backup.backup_error)

    def kill_excel(self):
        Excel_exit.execute_()
        self.ui.textBrowser.insertPlainText('Excel is dead\n')

    def add_to_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_5.text())) > 0:
            lists.cs_list_edit('1', self.ui.lineEdit_5.text())

    def delete_from_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_6.text())) > 0:
            lists.cs_list_edit('2', self.ui.lineEdit_6.text())
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText(lists.errors)

    def print_cs_cca_list(self):
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(', '.join(sorted(lists.cs_list)))

    def add_to_half_day_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_7.text())) > 0:
            lists.half_day_list_edit('1', self.ui.lineEdit_7.text())

    def delete_from_half_day_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_8.text())) > 0:
            lists.half_day_list_edit('2', self.ui.lineEdit_8.text())
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText(lists.errors)
            lists.errors = ''

    def print_half_day_cs_cca_list(self):
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(', '.join(sorted(lists.half_day_list)))

    def add_to_sales_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_9.text())) > 0:
            lists.sales_list_edit('1', self.ui.lineEdit_9.text())

    def delete_from_sales_CCA_cs_list(self):
        if len(str(self.ui.lineEdit_10.text())) > 0:
            lists.sales_list_edit('2', self.ui.lineEdit_10.text())
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText(lists.errors)
            lists.errors = ''

    def print_sales_cs_cca_list(self):
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(', '.join(sorted(lists.sales_list)))

    def update_cs_dict_(self):
        names = self.ui.lineEdit_12.text().split(',')
        try:
            lists.cs_dict_edit('1', names[0], names[1])
        except IndexError:
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText('Check the spelling, or field is empty\n')

    def add_to_cs_dict(self):
        names = self.ui.lineEdit_11.text().split(',')
        try:
            lists.cs_dict_edit('2', names[0], names[1])
        except IndexError:
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText('Check the spelling, or field is empty\n')

    def delete_from_cs_dict(self):
        if len(str(self.ui.lineEdit_13.text())) > 0:
            lists.cs_dict_edit('3', self.ui.lineEdit_13.text(), '0')
            self.ui.textBrowser.clear()
            self.ui.textBrowser.insertPlainText(lists.errors)
            lists.errors = ''

    def print_cs_dict(self):
        od = collections.OrderedDict(sorted(lists.cs_dict.items()))
        od = str(od)
        od = od.lstrip('OrderedDict')
        od = od.replace('(', '')
        od = od.replace('),', '\n')
        od = od.replace('[', '')
        od = od.replace(']', '')
        od = od.replace(')', '')
        self.ui.textBrowser.clear()
        self.ui.textBrowser.insertPlainText(od)

    def save_lists(self):
        lists.save()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())
