import gmail
import ot
import make_report
import lists
import config


Actions = {
    '1': gmail.execute_,
    '2': ot.execute_,
    '3': make_report.execute_,
    '4': lists.execute_,
    '5': config.change_paths
}
while True:
    print('''Choose action
    1 for downloading reports,
    2 for calculates ot in every file
    3 for make final report
    4 for editing workers lists
    5 for exit''')
    action = input('?')
    try:
        if action == '5':
            break
        Actions.get(action)()
    except KeyError:
        print('Incorrect Action')
