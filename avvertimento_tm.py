import os
import win32com.client as client
from PIL import ImageGrab
from directories import *


workbook_path = dashboard_inflow_canali_prodotti+r"\Condivisi\Dashboard inflow - Confronto SFDC vs. SAP.xlsb"

File = client.Dispatch('Excel.Application')
Workbook = File.Workbooks.Open(workbook_path)
Sheet = Workbook.Sheets.Item("Sintesi")
Copyrange= Sheet.Range('V8:Y13')
Copyrange.CopyPicture(Appearance=1, Format=2)

image_path = ImageGrab.grabclipboard().save(dashboard_inflow_canali_prodotti+r"\Condivisi\Avvertimento_tm.png")

Workbook.Close(SaveChanges=True)

File.Quit()

html_body = """
    <div>
        <p>Buongiorno a tutti,<br><br>
        vi riporto la situazione aggiornata a questa mattina (SF vs. SAP):<br><br></p>
    </div>
    <div>
        <img src={}></img>
    </div>
        <p>Appena vi è possibile per cortesia sistemate SF.<br><br>
        Un saluto,<br>Raffaele<br><br></p>
    </div>

"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Cristiano.Mozzolani@wolterskluwer.com;Andrea.Ferrara@wolterskluwer.com;rossella.specia@wolterskluwer.com;Gabriele.Albarello@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com;dimitri.falzone@wolterskluwer.com"
message.Subject = 'Salesforce vs SAP'
message.HTMLBody = html_body.format(image_path)

message.Display()
