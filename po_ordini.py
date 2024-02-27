import pandas as pd
import datetime as dt
import numpy as np
import zipfile
import shutil
import time
from directories import *


ins_man = shutil.copy(dashboard_inflow_prodotti + r"\inflow_prodotti_ins_man.xlsx",
            dashboard_inflow_source + r"\po_giornaliero\\")

zip_file = dashboard_inflow_canali_prodotti + r"\Data export\ZQ_PO_TAX_GIOR.zip"
directory_to_extract_to = dashboard_inflow_canali_prodotti + r"\Data export"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")



bw_file = dashboard_inflow_canali_prodotti + r"\Data export\ZQ_PO_TAX_GIOR_00000.xls"
df = pd.DataFrame(pd.read_html(str(bw_file), encoding = 'utf-8', decimal=",", thousands='.')[3])
po = pd.DataFrame(data=df.loc[1:,:])
po.columns = df.loc[0,:].to_list()


po.rename(columns={"Data Ins. Ordine":"Data Ins. Ordine (monitoraggio)","Imponibile Ordine":"Raccolto"}, inplace=True)
po["Budget  assegnato"] = np.nan
po["Data Ins. Ordine (monitoraggio)"] = pd.to_datetime(po["Data Ins. Ordine (monitoraggio)"], format="%d%m%Y")
po[["Raccolto", "Budget  assegnato"]] = po.loc[:,["Raccolto", "Budget  assegnato"]].astype(float)

po.to_excel(dashboard_inflow_source + r"\po_giornaliero\inflow_prodotti_po_giornaliero.xlsx", sheet_name='Foglio1', index=False)

po.groupby(["Numero Ordine","Data Ins. Ordine (monitoraggio)"], dropna=False).agg({"Raccolto":"sum"}).reset_index().to_excel(bpm_folder+r"\Ordini_SAP_PO.xlsx", sheet_name="Foglio1", index=False)



import win32com.client as client
from PIL import ImageGrab

workbook_path = dashboard_inflow_canali_prodotti + r"\Dashboard inflow - Raffaele - PO.xlsb"

File = client.Dispatch("Excel.Application")
Workbook = File.Workbooks.Open(workbook_path)
File.Visible = True
Workbook.RefreshAll()
File.CalculateUntilAsyncQueriesDone()
Workbook.Save()

shutil.move(dashboard_inflow_source + r"\po_giornaliero\inflow_prodotti_po_giornaliero.xlsx", 
            dashboard_inflow_source + r"\po_giornaliero_temp\inflow_prodotti_po_giornaliero.xlsx")


#Email SF vs. SAP
Sheet = Workbook.Sheets.Item("Email")
Copyrange= Sheet.Range('B5:I80')
Copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save(pictures + r'\Dashboard_Ordini.png')




image_path = dashboard_inflow_canali_prodotti + r'\Dashboard_Ordini.png'


html_body = """
    <div>
          <p>Marco,<br><br>
            qui sotto l'aggiornamento delle {ora}:<br><br>
                <div>
                    <img src={immagine}></img>
                </div><br><br>
            Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "marco.bitossi@wolterskluwer.com"
message.CC = "loredana.montagna@wolterskluwer.com;dimitri.falzone@wolterskluwer.com"
message.Subject = 'Aggiornamento Dashboard con dati del giorno'
message.HTMLBody = html_body.format(ora=dt.datetime.now().strftime("%H"),
                                    immagine=image_path)

message.Display()
