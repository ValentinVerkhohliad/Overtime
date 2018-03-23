import json


def execute():
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
                break
            actions.get(action)()
        except:
            print('Incorrect Action')


def load():
    try:
        with open('variables_list.json', 'rt') as f:
            return json.load(f)
    except FileNotFoundError:
        return []


def save():
    with open('variables_list.json', 'wt') as f:
        json.dump(variables_list, f)


variables_list = load()
cs_list, morning_workers_list, half_day_list, sales_list, cs_dict = (variables_list[0], variables_list[1],
                                                                     variables_list[2], variables_list[3],
                                                                     variables_list[4])


def cs_list_edit():
    global cs_list
    action = input('Please, choose action\n1.Add worker to CS list\n2.Delete worker from CS list\n3.Print list\n')
    if action == '1':
        cs_list.append(input('Please input worker CCA id'))
    elif action == '2':
        try:
            cs_list.remove(input('Please input worker CCA id'))
        except ValueError:
            raise ValueError('Worker is not in a list')
    elif action == '3':
        print(cs_list)


def morning_workers_list_edit():
    global morning_workers_list
    action = input('Please, choose action \n1. Add worker to  list\n2.Delete worker from list \n3. Print list')
    if action == '1':
        morning_workers_list.append(input('Please input worker CCA id'))
    elif action == '2':
        try:
            morning_workers_list.remove(input('Please input worker CCA id'))
        except ValueError:
            raise ValueError('Worker is not in a list')
    elif action == '3':
        print(morning_workers_list)


def half_day_list_edit():
    global half_day_list
    action = input('Please, choose action \n1. Add worker to  list \n2.Delete worker from list \n3. Print list')
    if action == '1':
        half_day_list.append(input('Please input worker CCA id'))
    elif action == '2':
        try:
            half_day_list.remove(input('Please input worker CCA id'))
        except ValueError:
            raise ValueError('Worker is not in a list')
    elif action == '3':
        print(half_day_list)


def sales_list_edit():
    global sales_list
    action = input('Please, choose action \n1. Add worker to  list \n2.Delete worker from list \n3. Print list')
    if action == '1':
        sales_list.append(input('Please input worker CCA id'))
    elif action == '2':
        try:
            sales_list.remove(input('Please input worker CCA id'))
        except ValueError:
            raise ValueError('Worker is not in a list')
    elif action == '3':
        print(sales_list)


def cs_dict_edit():
    global cs_dict
    action = input('Please, choose action \n 1.Update workers name \n 2.Add new worker to  list \n '
                   '3.Delete worker from list \n 4.Print list')
    if action == '1':
        cs_dict[input('Input new worker name')] = cs_dict.pop(input('Input old worker name to replace'))
    if action == '2':
        cs_dict[input('Input name of the new worker')] = input('Input CCA id of the new worker')
    if action == '3':
        del cs_dict[input('Input name of the worker')]
    elif action == '4':
        print(cs_dict)


actions = {
    '1': cs_list_edit,
    '2': morning_workers_list_edit,
    '3': half_day_list_edit,
    '4': sales_list_edit,
    '5': cs_dict_edit,
}