import win32com.client


def execute_():
    global xlApp
    xlApp = win32com.client.Dispatch("Excel.Application")
    xlApp.Quit()
