import win32com.client as client
from PIL import ImageGrab
from directories import *

workbook_path = dashboard_inflow_canali_prodotti+r"\Condivisi\Dashboard inflow - Confronto SFDC vs. SAP.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(workbook_path)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()


#Email SF vs. SAP
Sheet = Workbook.Sheets.Item("Sintesi")
Copyrange = Sheet.Range('L7:T14')
Copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(dashboard_inflow_canali_prodotti+r"\Condivisi\Dashboard inflow - Confronto SFDC vs. SAP.png')







html_body = """
    <div>
          <p>Buongiorno Marco,<br><br>
            qui il confronto SF vs. SAP:<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "marco.bitossi@wolterskluwer.com"
message.CC = "loredana.montagna@wolterskluwer.com"
message.Subject = 'Salesforce vs SAP'
message.HTMLBody = html_body

message.Display()
