import pandas as pd
import datetime as dt
import numpy as np
import zipfile
import shutil


ins_man = shutil.copy(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti\inflow_prodotti_ins_man.xlsx",
            r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\po_giornaliero\\")

zip_file =r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export\ZQ_PO_TAX_GIOR.zip"
directory_to_extract_to = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")



bw_file = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export\ZQ_PO_TAX_GIOR_00000.xls"
df = pd.DataFrame(pd.read_html(str(bw_file), encoding = 'utf-8', decimal=",", thousands='.')[3])
po = pd.DataFrame(data=df.loc[1:,:])
po.columns = df.loc[0,:].to_list()


po.rename(columns={"Data Ins. Ordine":"Data Ins. Ordine (monitoraggio)","Imponibile Ordine":"Raccolto"}, inplace=True)
po["Budget  assegnato"] = np.nan
po["Data Ins. Ordine (monitoraggio)"] = pd.to_datetime(po["Data Ins. Ordine (monitoraggio)"], format="%d%m%Y")
po[["Raccolto", "Budget  assegnato"]] = po.loc[:,["Raccolto", "Budget  assegnato"]].astype(float)

po.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\po_giornaliero\inflow_prodotti_po_giornaliero.xlsx", sheet_name='Foglio1', index=False)



import win32com.client as client
from PIL import ImageGrab

workbook_path = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Dashboard inflow - Raffaele - PO.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.open(workbook_path)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()

shutil.move(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\po_giornaliero\inflow_prodotti_po_giornaliero.xlsx",r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\po_giornaliero_temp\inflow_prodotti_po_giornaliero.xlsx")


#Email SF vs. SAP
excel = client.Dispatch('Excel.Application')
wb = excel.Workbooks.Open(workbook_path)
sheet = wb.Sheets.Item("vs. SF")
copyrange= sheet.Range('G7:J20')
copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\SFvsSAP.png')

image_path = r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\SFvsSAP.png'


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


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "marco.bitossi@wolterskluwer.com"
message.CC = "loredana.montagna@wolterskluwer.com"
message.Subject = 'Salesforce vs SAP'
message.HTMLBody = html_body.format(image_path)

message.Display()
