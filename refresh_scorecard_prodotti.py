import win32com.client as client
from directories import dashboard_inflow_canali_prodotti
import time 

############################################################################################################

scorecard_ind = dashboard_inflow_canali_prodotti + r"\Scorecard\Scorecard clienti su prodotti - Indiretta.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(scorecard_ind)
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


time.sleep(10)

html_body = """
    <div>
          <p>Ciao a tutti,<br><br>
            in allegato il file aggiornato.<br><br>
            Saluti,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Francesco.Sanseverino@wolterskluwer.com;francesca.panetta@wolterskluwer.com;Francesco.Florio@wolterskluwer.com;Giuseppe.Manuguerra@wolterskluwer.com;Elisa.Rigo@wolterskluwer.com;Rossella.Specia@wolterskluwer.com"
message.CC = "Camilla.Fabris@wolterskluwer.com;Sergio.Boaretto@wolterskluwer.com;Gianluca.Enea@wolterskluwer.com;Mirko.Fratus@wolterskluwer.com;PATRIZIA.SETTI@wolterskluwer.com;Massimiliano.Favoti@wolterskluwer.com; AlessandroAndrea.Travaglia@wolterskluwer.com;Cinzia.Borelli@wolterskluwer.com;Gianfranco.Altamore@wolterskluwer.com;Maurizio.Ferraresi@wolterskluwer.com;Dimitri.Falzone@wolterskluwer.com;Loredana.Montagna@wolterskluwer.com"
message.Subject = 'Scorecard clienti su prodotti - Indiretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=scorecard_ind)

message.Display()

#############################################################################################################

scorecard_dir = dashboard_inflow_canali_prodotti + r"\Scorecard\Scorecard clienti su prodotti - Diretta.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(scorecard_dir)
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


time.sleep(10)

html_body = """
    <div>
          <p>Ciao a tutti,<br><br>
            in allegato il file aggiornato.<br><br>
            Saluti,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "gabriele.albarello@wolterskluwer.com;Cristiano.Mozzolani@wolterskluwer.com;claudio.ferrante@wolterskluwer.com"
message.CC = "Gianluca.Enea@wolterskluwer.com;Sergio.Boaretto@wolterskluwer.com;Mirko.Fratus@wolterskluwer.com;PATRIZIA.SETTI@wolterskluwer.com;Massimiliano.Favoti@wolterskluwer.com; AlessandroAndrea.Travaglia@wolterskluwer.com;Cinzia.Borelli@wolterskluwer.com;Gianfranco.Altamore@wolterskluwer.com;Maurizio.Ferraresi@wolterskluwer.com;Dimitri.Falzone@wolterskluwer.com;Loredana.Montagna@wolterskluwer.com"
message.Subject = 'Scorecard clienti su prodotti - Diretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=scorecard_dir)

message.Display()

#############################################################################################################