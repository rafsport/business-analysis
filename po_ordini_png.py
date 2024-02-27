
from directories import *
import win32com.client as client
from PIL import ImageGrab

workbook_path = dashboard_inflow_canali_prodotti + r"\Dashboard inflow - Raffaele - PO.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(workbook_path)
File.Visible = True



#Email SF vs. SAP
Sheet = Workbook.Sheets.Item("Email")
Copyrange= Sheet.Range('B6:I74')
Copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(pictures + r'\Dashboard_Ordini.png')