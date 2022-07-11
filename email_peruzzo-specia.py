import win32com.client as client
#############################################################################################################
# Specia
#
dashboard_inflow_indiretta = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Condivisi\Dashboard inflow - Indiretta.xlsb"

html_body = """
    <div>
          <p>Ciao Rossella,<br><br>
            in allegato il file aggiornato.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "rossella.specia@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com;claudio.ferrante@wolterskluwer.com"
message.Subject = 'Inflow indiretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=dashboard_inflow_indiretta)

message.Display()


############################################################################################################
#Peruzzo

dashboard_inflow_clienti_inside = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\Dashboard inflow - Clienti inside sales.xlsb"

image_path = r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\clienti_inside.png'


html_body = """
    <div>
          <p>Ciao Laura,<br><br>
            in allegato il file con i dati aggiornati.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
    <div>
        <img src={}></img>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Laura.Peruzzo@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com;claudio.ferrante@wolterskluwer.com"
message.Subject = 'Aggiornamento inflow clienti inside sales'
message.HTMLBody = html_body.format(image_path)
message.Attachments.Add(Source=dashboard_inflow_clienti_inside)

message.Display()
