import gmail
import ot
import make_report
import lists
Actions = {
    '1': gmail.exec_,
    '2': ot.execute_,
    '3': make_report.execut_,
    '4': lists.execute,
}
while True:
    print('''Choose action
    1 for downloading reports,
    2 for calculates ot in every file
    3 for make final report
    4 for editing workers lists
    5 for exit''')
    action = input('!')
    try:
        if action == '5':
            break
        Actions.get(action)()
    except:
        print('Incorrect Action')
