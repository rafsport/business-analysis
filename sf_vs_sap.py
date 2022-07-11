import win32com.client as client

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
