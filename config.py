import json


def change_paths():
    while True:
        action=input('Please, choose action \n1.Change main report Path \n2.Change downloading reports path \n3.Exit\n')
        if action == '1':
            change_mrp()
        if action == '2':
            change_rp()
        if action == '3':
            save()
            break


def load():
    try:
        with open('config.json', 'rt') as f:
            return json.load(f)
    except FileNotFoundError:
        return ['', '']


def save():
    config = [change_mrp(), change_rp()]
    with open('config.json', 'wt') as f:
        json.dump(config, f)


def reports_path(rep_path='c:\\reports\\logins\\'):
    return rep_path + '*xlsx'


def main_rep_path_(main_rep_path="C:\\reports\\reports\\Ot.xlsx"):
    return main_rep_path


def change_mrp():
    return main_rep_path_(input('Input new path'))


def change_rp():
    return reports_path(input('Input new path'))



