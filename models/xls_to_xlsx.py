import os

import win32com.client as client


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
