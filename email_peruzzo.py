import win32com.client as client

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

############################################################################################################
#Disimino

image_path = r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\clienti_inside.png'


html_body = """
    <div>
          <p>Ciao Mattia,<br><br>
            in allegato il file png con i dati aggiornati.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
    <div>
        <img src={}></img>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Mattia.Disimino@wolterskluwer.com"
message.Subject = 'Picture inflow clienti inside sales'
message.HTMLBody = html_body.format(image_path)
message.Attachments.Add(Source=image_path)

message.Display()
