import win32com.client as client
import os

# Apro il file Excel
excel = client.Dispatch("Excel.Application")
workbook = excel.Workbooks.Open(os.path.abspath(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Dashboard inflow - Raffaele.xlsb"))

# Ottengo le tabelle pivot dal foglio di lavoro desiderato
worksheet = workbook.Worksheets("Email")
pivot_tables = worksheet.PivotTables

# Creo una nuova email
outlook = client.Dispatch("Outlook.Application")
message = outlook.CreateItem(0)
message.Subject = "Oggetto dell'email"
message.To = "raffaele.sportiello@wolterskluwer.com"

# Aggiungo le tabelle pivot come immagini nel corpo dell'email
for i in range(1, pivot_tables.Count+1):
    pivot_table = pivot_tables(i)
    pivot_table.Chart.Export(os.path.abspath(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\image.png"), "png")
    attachment = message.Attachments.Add(os.path.abspath(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\image.png"))
    cid = "image{}".format(i)
    image_tag = '<img src="cid:{}">'.format(cid)
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
    message.HTMLBody += image_tag

# Invio l'email
message.Display()

# Chiudo il file Excel
workbook.Close(False)
excel.Quit()