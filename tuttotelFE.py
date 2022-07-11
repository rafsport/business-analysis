import os
import win32com.client as client
from PIL import ImageGrab


workbook_path = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Dashboard inflow - Raffaele.xlsb"

excel = client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(workbook_path)
sheet = wb.Sheets.Item("TuttotelFE")
copyrange= sheet.Range('B5:G30')
copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\TuttotelFE.png')

image_path = r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\TuttotelFE.png'

html_body = """
    <div>
          <p>Buongiorno Claudio,<br><br>
            riporto in calce l’aggiornamento sui dati di inflow di Tuttotel FE.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
    <div>
        <img src={}></img>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "claudio.ferrante@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com"
message.Subject = 'Tuttotel FE'
message.HTMLBody = html_body.format(image_path)

message.Display()

#message.Send()
