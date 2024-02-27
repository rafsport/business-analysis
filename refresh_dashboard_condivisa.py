import win32com.client

File = win32com.client.Dispatch("Excel.Application")
Workbook = File.Workbooks.open(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Condivisi\Dashboard inflow.xlsb")
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()
File.Quit()
