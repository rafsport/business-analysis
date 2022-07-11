import win32com.client

File = win32com.client.Dispatch("Excel.Application")
Workbook = File.Workbooks.open(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Dashboard inflow - Raffaele.xlsb")
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()


File = win32com.client.Dispatch("Excel.Application")
Workbook = File.Workbooks.open(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Condivisi\Dashboard inflow.xlsb")
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()
File.Quit()




html_body = """
    <div>
          <p>Buongiorno Marco,<br><br>
            qui il confronto SF vs. SAP:<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""

outlook = win32com.client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "marco.bitossi@wolterskluwer.com"
message.CC = "loredana.montagna@wolterskluwer.com"
message.Subject = 'Salesforce vs SAP'
message.HTMLBody = html_body

message.Display()
