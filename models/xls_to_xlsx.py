import os

import script
import win32com.client as client

""" Converts the .xls files into .xlsm files """
def converter():
    excel = client.Dispatch("excel.application")
    # Converts the files from xls to xlsx
    for file in os.listdir(os.getcwd() + "/old version/"):
        filename, fileextenstion = os.path.splitext(file)
        wb = excel.Workbooks.Open(os.getcwd() + "/old version/" + file)
        output_path = os.getcwd() + "/new version/" + filename
        wb.SaveAs(output_path, 51)
        wb.Close()
    excel.Quit()

""" Searches A column for Headers """
def search_value_in_column(ws, search_string, column='A'):
    for row in range(1, ws.max_row + 1):
        coordinate = "{}{}".format(column, row)
        if ws[coordinate].value == search_string:
            return column, row
    return column, None

""" Searches Row for total """
def search_value_in_row(ws, search_string, row):
    for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H','I','K']:
        coordinate = "{}{}".format(column, row)
        if ws[coordinate].value == search_string:
            return column, row
    return row, None

""" Checks which file it currently is scrapping from """
def check_file(coordinate, value):
    if script.files == script.file_list[0]:
            script.main_sheet['B' + coordinate] = value
    elif script.files == script.file_list[1]:
        script.main_sheet['C' + coordinate] = value
    elif script.files == script.file_list[2]:
        script.main_sheet['D' + coordinate] = value
    elif script.files == script.file_list[3]:
        script.main_sheet['E' + coordinate] = value
    elif script.files == script.file_list[4]:
        script.main_sheet['F' + coordinate] = value
    elif script.files == script.file_list[5]:
        script.main_sheet['G' + coordinate] = value
    elif script.files == script.file_list[6]:
        script.main_sheet['H' + coordinate] = value
