import win32com.client as client
from directories import dashboard_inflow_canali_prodotti
import time 

############################################################################################################

dashboard_inflow_indiretta = dashboard_inflow_canali_prodotti + r"\Condivisi\Dashboard inflow - Indiretta.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(dashboard_inflow_indiretta)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()
File.Quit()

import psutil

def kill_excel():
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()

#############################################################################################################
time.sleep(10)

html_body = """
    <div>
          <p>Ciao Francesco,<br><br>
            in allegato il file aggiornato.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Francesco.Sanseverino@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com;dimitri.falzone@wolterskluwer.com;claudio.ferrante@wolterskluwer.com;camilla.fabris@wolterskluwer.com;Deroghe.GTM@wolterskluwer.com;trademarketing-IT@wolterskluwer.com"
message.Subject = 'Inflow indiretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=dashboard_inflow_indiretta)

message.Display()
