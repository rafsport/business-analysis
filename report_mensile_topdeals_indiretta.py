import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join
import calendar
idx = pd.IndexSlice


inflow_prodotti = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti\\"

onlyfiles = [f for f in listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti") if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti", f))]

df = pd.DataFrame()
for file in onlyfiles:
    data = pd.read_excel(inflow_prodotti+file)
    df = df.append(data)

for col in df.select_dtypes(include=[object]).columns:
    df[col] = df[col].astype(str)


## Parametriche

parametriche = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche\\"
onlyfiles = [f for f in listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche") if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche", f))]

par = {}
for file in onlyfiles:
    file_name=file.split(".")[0]
    par[f"{file_name}"] = pd.read_excel(parametriche+file)


## All accounts da Salesforce

accounts_sf = pd.read_csv(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Salesforce\Data export\IT_all_accounts.csv", encoding='latin-1', sep=";")

# Deals

filt_rows = (df["Data Ins. Ordine (monitoraggio)"].dt.year == 2022)&(df["Raccolto"].notna())
columns = ["Cliente Merce", "Agente dell'ORDINE", "Sales Director", "Canale di VENDITA", "Codice PRODOTTO", "Linea Prodotti", "Numero Ordine", "Numero Fattura", "Data Ins. Ordine (monitoraggio)", "Raccolto"]

df_2022 = df.loc[filt_rows, columns].reset_index(drop=True)
df_2022["Sales Director"] = df_2022["Sales Director"].str.lower()
df_2022["Agente dell'ORDINE"] = df_2022["Agente dell'ORDINE"].str.lower()

esc_fatture = par["Esclusione_fatture"]
esc_fatture["Numero Fattura"] = esc_fatture["Numero Fattura"].astype(str)

esc_ordini = par["Esclusione_ordini"]
esc_ordini["Numero Ordine"] = esc_ordini["Numero Ordine"].astype(str)

sales_director = par["Sales_director"]
sales_director["Sales Director"] = sales_director.loc[:,"Sales Director"].str.lower()
sales_director = sales_director[["Sales Director", "Rete", "Area territoriale"]]

agente = par["Agente_ordine"]
agente["Codice agente"] = agente.loc[:,"Codice agente"].str.lower()

cod_prod = par["Classificazione_codice_prodotto"]
cod_prod["Codice PRODOTTO"] = cod_prod["Codice PRODOTTO"].apply(lambda x: str(x))
cod_prod = cod_prod.loc[:, ['MDM', 'GTM BDG', 'Solution','Sub solution','Descrizione prodotto',"Codice PRODOTTO"]]

li_prodotti = par["Linea_prodotti"]

inflow = df_2022.merge(esc_fatture, on="Numero Fattura", how="left").merge(esc_ordini, on="Numero Ordine", how="left").merge(li_prodotti, on="Linea Prodotti", how="left").merge(sales_director, on="Sales Director", how="left").merge(agente, how="left", left_on="Agente dell'ORDINE", right_on="Codice agente").merge(cod_prod, how="left", on="Codice PRODOTTO")

esclusioni = (inflow["Canale di VENDITA"] != "E-Commerce")&(inflow["Escludi fatture"].isna())&(inflow["Escludi ordine"].isna())

inflow = inflow.loc[esclusioni, ['Data Ins. Ordine (monitoraggio)', "Cliente Merce", "Numero Ordine", "Rete", "Area territoriale", 'Sales Director', "Agente dell'ORDINE_y","RSM Agente",'Linea', "MDM","Solution","Descrizione prodotto", 'Raccolto']].reset_index(drop=True)

inflow.rename(columns={"Data Ins. Ordine (monitoraggio)":'Giorno', "Agente dell'ORDINE_y":"Agenzia"}, inplace=True)

inflow["Cliente Merce"] = inflow["Cliente Merce"].astype(str).apply(lambda x: "IT-" + x.zfill(10))
inflow = inflow.merge(accounts_sf, how="left", left_on="Cliente Merce", right_on="WK Account Number")

filt = (inflow["Rete"].str.contains("Indiretta")) & (inflow["Giorno"].dt.date >= dt.date(dt.datetime.now().year,dt.datetime.now().month-1, 1)) & (inflow["Giorno"].dt.date <= dt.date(dt.datetime.now().year,dt.datetime.now().month-1, calendar.monthrange(dt.datetime.now().year,dt.datetime.now().month-1)[1]))
columns = ["Giorno","Numero Ordine","Cliente Merce","Account Name","Agenzia","Area territoriale","Rete","Linea","Solution","Raccolto"]

inflow_ind = inflow.loc[filt, columns]

inflow_ind["Valore Ordine"] = inflow_ind.groupby(["Numero Ordine"])["Raccolto"].transform("sum")

topdeals_ind = inflow_ind.groupby(["Numero Ordine","Account Name", "Cliente Merce", "Agenzia", "Area territoriale","Solution", inflow_ind["Giorno"]]).agg({"Raccolto":"sum", "Valore Ordine":"max"}).sort_values(by="Valore Ordine", ascending=False).reset_index()
#.style.format({"Raccolto":"{:,.0f}"})

topdeals_ind.rename(columns={"Account Name":"Nome cliente", "Cliente Merce":"Codice SAP", "Solution":"Prodotto", "Raccolto":"Inflow"}, inplace=True)
topdeals_ind = topdeals_ind.loc[topdeals_ind["Valore Ordine"] > 0, :].copy()

with pd.ExcelWriter(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Speciali\Top deals indiretta.xlsx", engine="openpyxl") as writer:
    topdeals_ind.to_excel(writer, sheet_name="topdeals_ind", index=False)


import win32com.client as client
#############################################################################################################
# Specia
#
dashboard_inflow_indiretta = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Condivisi\Dashboard inflow - Indiretta.xlsb"

html_body = """
    <div>
          <p>Ciao Rossella,<br><br>
            in allegato trovi il file con l’elenco dei deal dell’indiretta del mese appena concluso.<br><br>
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
