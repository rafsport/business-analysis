import win32com.client
from PIL import ImageGrab
import time
from directories import *
import datetime as dt

File = win32com.client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(dashboard_inflow_canali_prodotti+r"\Dashboard inflow - Raffaele.xlsb")
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()


Worksheet = Workbook.Worksheets("Email")

Copyrange= Worksheet.Range('B6:k102')
Copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(pictures+r"\Dashboard inflow.png")

Workbook.Close()


time.sleep(5)




Workbook = File.Workbooks.Open(dashboard_inflow_canali_prodotti+r"\Condivisi\Dashboard inflow.xlsb")
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()
Workbook.Close(SaveChanges=False)



time.sleep(5)



workbook_path = dashboard_inflow_canali_prodotti+r"\Condivisi\Dashboard inflow - Confronto SFDC vs. SAP.xlsb"


Workbook = File.Workbooks.Open(workbook_path)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()


#Email SF vs. SAP
Sheet = Workbook.Sheets.Item("Sintesi")
Copyrange = Sheet.Range('L7:T14')
Copyrange.CopyPicture(Appearance=1, Format=2)

image_path = ImageGrab.grabclipboard().save(pictures+r"\Dashboard inflow - Confronto SFDC vs. SAP.png")

Workbook.Close(SaveChanges=False)

File.Quit()


html_body = """
    <div>
          <p>Buongiorno Marco,<br><br>
            qui il confronto SF vs. SAP:<br><br>
            <div>
                <img src={}></img>
            </div><br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""

outlook = win32com.client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "marco.bitossi@wolterskluwer.com"
message.CC = "loredana.montagna@wolterskluwer.com;dimitri.falzone@wolterskluwer.com"
message.Subject = 'Salesforce vs SAP'
message.HTMLBody = html_body.format(image_path)

message.Display()



###############################################################
#<div><img width="280" height="45" src=https://cdn.wolterskluwer.io/wk/fundamentals/1.15.2/logo/assets/medium.svg></img></div>

html_body = """

    <div>
        <p>
        Buongiorno,<br><br>
        in allegato la dashboard aggiornata.<br><br>
        Raffaele
        </p>


        <div><hr size="1" width="100%" align="center"></div>
        <div><hr size="1" width="100%" align="center"></div>
        
        
        <br>
        <div><p>Aggiornamento al {}</p></div>

        <div>
        <span lang="IT">
            <a href="https://nam04.safelinks.protection.outlook.com/ap/x-59584e83/?url=https%3A%2F%2Fwolterskluwer-my.sharepoint.com%2F%3Ax%3A%2Fp%2Fraffaele_sportiello%2FEaIMGC62qmZFruYFdZFtpqUBF4RZNhrybnsO_-vTklLzTA%3Fe%3DiPyaXe&amp;data=05%7C01%7CRaffaele.Sportiello%40wolterskluwer.com%7C7656ec3191b7425e5ddf08db1e13f477%7C8ac76c91e7f141ffa89c3553b2da2c17%7C0%7C0%7C638136843317728776%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&amp;sdata=HzXyl5BWjxnXomNvjcL5qEAszwqkK%2Bo1Wp%2BHnJ3zTQY%3D&amp;reserved=0" originalsrc="https://wolterskluwer-my.sharepoint.com/:x:/p/raffaele_sportiello/EaIMGC62qmZFruYFdZFtpqUBF4RZNhrybnsO_-vTklLzTA?e=iPyaXe" shash="zaNwpHtkDDyF+cayCzaAtsK8lAqNhyt2033mQ5Y9QuS6seLu4pYbnL4SJk2JdGJ8amY+ftr1dBEzuE0R9teOdxfsbcp8U055kRtDzMbfFgQpitxl6qzXv2mHgKvRNPnKmox3RXw4xsIwrJVtvlnAlNGlZ51k4/Pzau8clAs8GFU=">
            <span lang="IT">Clicca qui per visualizzare la dashboard completa</span>
            </a>
        </span>

        <p style="  background:#007AC3;font-size:18.0pt;color:white;letter-spacing:.75pt;font-weight:bold;">
            <span lang="IT">Inflow</span>
            <span lang="IT">mese </span>
        </p>
    
        <div>
        Dashboard
        </div>
        <div>
            
        </div>

    
    </div>
"""

outlook = win32com.client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "raffaele.sportiello@wolterskluwer.com"
message.Subject = 'Dashboard inflow'
message.HTMLBody = html_body.format((dt.date.today()-dt.timedelta(days=1)).strftime('%d/%m/%Y'))

message.Display()

