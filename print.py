import win32com.client as win32


def print_pdf(file_to_print):

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(file_to_print)
    wb.PrintOut()

    wb.Close(False)
    excel.Quit()

    print("Impression terminée.")
