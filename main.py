#! python3

import os
import subprocess
import openpyxl


def createFolder(path):
    ''' creates folder if it doesn't already exist '''
    try:
        if not os.path.exists(path):
            os.mkdir(path)
    except OSError:
        print('Error creating directory' + path)


def open_excel(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path, data_only=True) # , data_only=True because scheduling file has a lot of formulas
        ws = wb.active
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
    wb.save(excel_file_path)
    return wb, ws


createFolder('.\\AUMID Excel')

# excels = os.listdir('.\\AUMID Excel')
# print(excels)
wb, ws = open_excel('.\\AUMID Excel\AUMID Excel.xlsx')
# wb = openpyxl.Workbook()
# ws = wb.active
bla = subprocess.run('"C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" get-StartApps',capture_output=True, text=True, shell=True)
print(bla)
print(bla.stdout)
ws.cell(row=1, column=1).value = bla.stdout
wb.save('.\\AUMID Excel\AUMID Excel.xlsx')
