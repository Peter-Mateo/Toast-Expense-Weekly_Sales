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