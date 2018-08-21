import gmail
import ot
import make_report
import lists
import work_days
import backup


Actions = {
    '1': gmail.execute_,
    '2': ot.execute_,
    '3': make_report.execute_,
    '4': work_days.execute_,
    '5': backup.execute_,
    '6': lists.execute_,
}
while True:
    print('''Choose action
    1 for downloading reports,
    2 for calculates ot in every file
    3 for make final report
    4 for fill apsend file
    5 for making backups
    6 for editing workers lists
    7 for exit''')
    action = input('?')
    try:
        if action == '6':
            break
        Actions.get(action)()
    except KeyError:
        print('Incorrect Action')
