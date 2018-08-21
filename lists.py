import json
errors = ''


def execute_():
    global errors
    while True:
        print('''Choose action
        1 for editing global cs worker list(only CCA id's),
        2 for editing morning workers list(only CCA id's)
        3 for editing half day workers list(only CCA id's)
        4 for editing sales workers list(only CCA id's)
        5 for editing cs workers list(names and id's)
        6 for exit''')
        action = input('?')
        try:
            if action == '6':
                save()
                break
            actions.get(action)()
        except KeyError:
            errors = errors + 'Incorrect Action'
        except ValueError as e:
            print(e)


def action_(fn):
    def dec():
        res = fn(action=input('Please, choose action \n1. Add worker to  list\n2.Delete worker from list \n3.'
                              ' Print list\n'
                              ))
        return res
    return dec


def load():
    try:
        with open('variables_list.json', 'rt') as f:
            return json.load(f)
    except FileNotFoundError:
        return []


def save():
    with open('variables_list.json', 'wt') as f:
        json.dump(variables_list, f)


def cs_list_edit(action, cca_number):
    global cs_list
    global errors
    if action == '1':
        cs_list.append(cca_number)
    elif action == '2':
        try:
            cs_list.remove(cca_number)
        except ValueError:
            errors = errors + 'Worker is not in a list\n'
    elif action == '3':
        return cs_list


@action_
def morning_workers_list_edit(action):
    global morning_workers_list
    global errors
    if action == '1':
        morning_workers_list.append(input('Please input worker CCA id'))
    elif action == '2':
        try:
            morning_workers_list.remove(input('Please input worker CCA id'))
        except ValueError:
            errors = errors + 'Worker is not in a list\n'
    elif action == '3':
        print(morning_workers_list)


def half_day_list_edit(action, cca_number):
    global half_day_list
    global errors
    if action == '1':
        half_day_list.append(cca_number)
    elif action == '2':
        try:
            half_day_list.remove(cca_number)
        except ValueError:
            errors = errors + 'Worker is not in a list\n'
    elif action == '3':
        return half_day_list


def sales_list_edit(action, cca_number):
    global sales_list
    global errors
    if action == '1':
        sales_list.append(cca_number)
    elif action == '2':
        try:
            sales_list.remove(cca_number)
        except ValueError:
            errors = errors + 'Worker is not in a list\n'
    elif action == '3':
        return sales_list


def cs_dict_edit(action, name, cca_number):
    global cs_dict
    global errors
    try:
        if action == '1':
            cs_dict[name] = cs_dict.pop(cca_number)
        if action == '2':
            cs_dict[name] = cca_number
        if action == '3':
            del cs_dict[name]
        elif action == '4':
            return cs_dict
    except KeyError:
        errors = errors + 'Worker is not in a list\n'
    except TypeError:
        errors = errors + 'Check the spelling\n'


actions = {
    '1': cs_list_edit,
    '2': morning_workers_list_edit,
    '3': half_day_list_edit,
    '4': sales_list_edit,
    '5': cs_dict_edit,
}

variables_list = load()
cs_list, morning_workers_list, half_day_list, sales_list, cs_dict = (variables_list[0], variables_list[1],
                                                                     variables_list[2], variables_list[3],
                                                                     variables_list[4])
