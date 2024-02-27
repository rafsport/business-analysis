import os
import win32com.client as client
from PIL import ImageGrab

############################################################################################################
# Peruzzo
dashboard_inflow_clienti_inside = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\Dashboard inflow - Clienti inside sales.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.open(dashboard_inflow_clienti_inside)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()

Sheet = Workbook.Sheets.Item("Forecast Call")
Copyrange = Sheet.Range('B7:O27')
Copyrange.CopyPicture(Appearance=1, Format=2)

Workbook.Save()
File.Quit()

ImageGrab.grabclipboard().save(r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\clienti_inside.png')
