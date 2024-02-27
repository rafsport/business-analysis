import win32com.client as client
#############################################################################################################
# Specia
#
dashboard_inflow_indiretta = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Condivisi\Dashboard inflow - Indiretta.xlsb"

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
message.CC = "marco.bitossi@wolterskluwer.com;claudio.ferrante@wolterskluwer.com;camilla.fabris@wolterskluwer.com;Deroghe.GTM@wolterskluwer.com;trademarketing-IT@wolterskluwer.com"
message.Subject = 'Inflow indiretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=dashboard_inflow_indiretta)

message.Display()
