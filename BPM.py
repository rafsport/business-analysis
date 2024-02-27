import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join
from os.path import getmtime
from directories import *
import time
#import po_ordini

data_agg = dt.datetime.fromtimestamp(getmtime(bpm_export_file))

# date_parser=lambda x: dt.datetime.strptime(x, '%d/%m/%Y')
bpm = pd.read_csv(bpm_export_file, sep=";", parse_dates=["Data Caricamento (in BPM)"], date_format='%d/%m/%Y', decimal=",", on_bad_lines="skip")

par = {}
for file in parametriche_files:
    file_name=file.split(".")[0]
    par[f"{file_name}"] = pd.read_excel(parametriche+"\\"+file)


agente = pd.DataFrame(par["Agente_ordine"][["Codice agente", "Agente dell'ORDINE", "RSM Agente", "Sales Director"]])

sales_director = pd.DataFrame(par["Sales_director"]).loc[:,["Sales Director", "Area territoriale", "Rete"]]

#sales_director = pd.read_excel(par_file_path["sales_director"], usecols=["Sales Director", "Area territoriale", "Rete"])

columns = ['Data Caricamento (in BPM)',  'Stato Lavorazione','Cod. Pratica (BPM)', 'Nr. Ordine','Inserito da', 'Sales Office', 'Codice Cliente SAP', 'Partita Iva/CF','Nome Cliente', 'Data Documento', 'Ordine Omaggio','Valore Totale Ordine', 'Importo HW', 'Importo Licenza','Importo Servizi', 'Note', 'Ord. Apparound','Automatico SAP', 'Opportunity Id', "Sales Director", "Rete", "Area territoriale", "Agente dell'ORDINE", "RSM Agente"]

# Confronto su ordini SAP
dfs = []
for file in [f for f in listdir(bpm_folder) if isfile(join(bpm_folder, f)) and "SAP" in f]:
    data = pd.read_excel(bpm_folder+"\\"+file, na_values="#", dtype={"Numero Ordine":str})
    data["Origine"] = file.split("_")[1].split(".")[0]
    data["Ultimo Agg."] = dt.date.fromtimestamp(getmtime(bpm_folder+"\\"+file))
    dfs.append(data)
ordini_sap = pd.concat(dfs, axis=0, ignore_index=True)

ordini_sap.to_excel(bpm_folder+r"\Ordini_totale.xlsx", index=False)

try:
    ordini_trovati = bpm.loc[bpm["Nr. Ordine"].isin(ordini_sap["Numero Ordine"]),"Nr. Ordine"].unique()
    if len(ordini_trovati) >0:
        print(f"I seguenti ordini sono già presenti in SAP:{ordini_trovati}")
        bpm = bpm.loc[~bpm["Nr. Ordine"].isin(ordini_trovati),:]
except:
    pass



bpm_def = ((bpm.loc[:,:].merge(agente, how="left",left_on="Sales Office", right_on="Codice agente"))\
    .merge(sales_director, left_on="Sales Director_y", right_on="Sales Director", how="left"))\
        .loc[:,columns]

bpm_gb_sintesi = bpm_def.groupby(["Rete", "Area territoriale"], dropna=False).agg({"Valore Totale Ordine":"sum", "Importo Licenza":"sum", "Importo Servizi":"sum", "Importo HW":"sum"}).assign(Totale_Importi=lambda df_: df_[["Importo Licenza", "Importo Servizi", "Importo HW"]].sum(axis=1))

bpm_gb_sintesi_totale = bpm_gb_sintesi.agg("sum").rename("Totale").to_frame().T
bpm_gb_sintesi_totale.index = pd.MultiIndex.from_tuples([("Totale", '')], names=('Rete', 'Area territoriale'))
bpm_gb_sintesi = pd.concat([bpm_gb_sintesi, bpm_gb_sintesi_totale], axis=0)

bpm_gb_numero_ord = bpm_def.groupby(["Rete", "Area territoriale", "Data Caricamento (in BPM)", "Inserito da", "Nr. Ordine", "Stato Lavorazione"], dropna=False).agg({"Valore Totale Ordine":"sum", "Importo Licenza":"sum", "Importo Servizi":"sum", "Importo HW":"sum"})


registro_ordini = bpm_folder+r"\Registro ordini inseriti in BPM.xlsx"
with pd.ExcelWriter(registro_ordini, datetime_format='YYYY-MM-DD') as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet('Ordini caricati')
    worksheet.insert_image('A1', pictures + r'\Dashboard_Ordini.png')
    bpm_gb_sintesi.to_excel(writer, sheet_name="Sintesi", float_format="%.0f")
    #bpm_gb_numero_ord.to_excel(writer, sheet_name="Numero ord")
    bpm_def.to_excel(writer, sheet_name="Dettaglio", index=False)

#bpm_gb_numero_ord.to_excel(writer, sheet_name="Numero ord", float_format="%.0f")
#############################################################################################################
# Format rows

import openpyxl

# Apri il file Excel con openpyxl
workbook = openpyxl.load_workbook(bpm_folder+r"\Registro ordini inseriti in BPM.xlsx")
worksheet = workbook['Sintesi']

worksheet["I1"] = "Aggiornamento al:"
worksheet["J1"] = data_agg

# Cerca la riga contenente "Totale" e applica lo stile "bold"
for row in worksheet.iter_rows():
    for cell in row:
        if cell.value == 'Totale':
            for c in row:
                c.font = openpyxl.styles.Font(bold=True)


for col in worksheet.columns:
    for cell in col:
        if cell.row > 1:
            cell.number_format = '#,##0'


# Salva il file Excel modificato
workbook.save(bpm_folder+r"\Registro ordini inseriti in BPM.xlsx")

#############################################################################################################
# Autofit columns

import win32com.client as client
excel = client.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(registro_ordini)

for sheet in wb.Sheets:
      
    ws = wb.Worksheets(sheet.Name)
    ws.Columns.AutoFit()

wb.Save()
excel.Application.Quit()


time.sleep(5)


#############################################################################################################
# Email
#
image_path = pictures + r'\Dashboard_Ordini.png'
attachments = [registro_ordini, pictures + r'\Dashboard_Ordini.png']

html_body = """
    <div>
          <p>Buonasera,<br><br>
            di seguito la situazione aggiornata alle ore {ora}:
            <br><br>
            <ul>
            <li>Valore ordini in SAP:  k</li>
            <li>Valore ordini in BPM:  k</li>
            </ul>
            <br>
            Riporto qui di seguito il dettaglio degli ordini presenti in SAP:<br><br>
            <div>
                <img src={immagine}></img>
            </div><br><br>
            <br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Pierfrancesco.Angeleri@wolterskluwer.com;marco.bitossi@wolterskluwer.com;claudio.ferrante@wolterskluwer.com;Susanna.Fontana@wolterskluwer.com;Simona.Sorbello@wolterskluwer.com;Camilla.Fabris@wolterskluwer.com;Andrea.Ferrara@wolterskluwer.com;Gabriele.Albarello@wolterskluwer.com;Cristiano.Mozzolani@wolterskluwer.com;Laura.Peruzzo@wolterskluwer.com;Francesca.Pepe@wolterskluwer.com"
message.CC = "dimitri.falzone@wolterskluwer.com;Loredana.Montagna@wolterskluwer.com;"
message.Subject = 'Aggiornamento ordini BPM'
message.HTMLBody = html_body.format(ora=dt.datetime.now().strftime("%H"),
                                    immagine=image_path)

for attachment in attachments:
    message.Attachments.Add(Source=attachment)

message.Display()


#message.To = "Cristiano.Mozzolani@wolterskluwer.com;Andrea.Ferrara@wolterskluwer.com;rossella.specia@wolterskluwer.com;Gabriele.Albarello@wolterskluwer.com;Camilla.Fabris@wolterskluwer.com"
#message.CC = "marco.bitossi@wolterskluwer.com;Loredana.Montagna@wolterskluwer.com;Susanna.Fontana@wolterskluwer.com;Alessia.Berra@wolterskluwer.com;Claudio.Ferrante@wolterskluwer.com;dimitri.falzone@wolterskluwer.com"
