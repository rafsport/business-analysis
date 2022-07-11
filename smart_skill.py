#!/usr/bin/env python
# coding: utf-8

# In[21]:


import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join
import zipfile



# Codici prodotto

cod_prodotti = pd.read_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche\Classificazione_codice_prodotto.xlsx")
cod_prodotti["Codice PRODOTTO"] = cod_prodotti["Codice PRODOTTO"].apply(lambda x:  x.zfill(8))
cod_prodotti.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data source\Classificazione_codice_prodotto.xlsx")



# Acquisti (report BUDEST)

zip_file =r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data export\ZQ_MC_BUDANAL_IMP_ALL.ZIP"
directory_to_extract_to = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data export"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")

bw_file = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data export\ZQ_MC_BUDANAL_IMP_ALL_00000.xls"
df = pd.DataFrame(pd.read_html(str(bw_file), encoding = 'utf-8', decimal=",", thousands='.')[3])
df2 = pd.DataFrame(data=df.loc[1:,:])
df2.columns = df.loc[0,:].to_list()

df2["Data Ins. Ordine (monitoraggio)"] = pd.to_datetime(df2["Data Ins. Ordine (monitoraggio)"], format="%d%m%Y")
df2["Codice PRODOTTO"] = df2["Codice PRODOTTO"].astype(str).apply(lambda x:  x.zfill(8))
df2["Cliente FATTURA"] = df2["Cliente FATTURA"].astype(str)
df2["Raccolto"] = df2["Raccolto"].astype(float)
aggiornamento = str(df2["Data Ins. Ordine (monitoraggio)"].max())[0:10]

onlyfiles = [f for f in listdir("C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/Ordini") if isfile(join("C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/Ordini", f))]

lastFile = onlyfiles[-1]
lastDatetime = pd.to_datetime(lastFile.split("_")[1].split(".")[0], format='%Y-%m-%d') + pd.Timedelta(1, unit="Day")

df2 = df2.loc[df2["Data Ins. Ordine (monitoraggio)"] >=lastDatetime]

df2.to_excel(f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer\Documents/E-commerce/Data source/Ordini/ordini_{aggiornamento}.xlsx", index=False)



# Report Shop Fattura Smart Orders

import xml.etree.ElementTree as ET

xml_file = f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data export/ReportShopFattSmartOrders/ReportShopFattSmartOrders.xls"

with open(xml_file) as fp:
    content = fp.read()
    content = content.replace('&', '&amp;')
    xml = ET.ElementTree(ET.fromstring(content))

L = []
rows = xml.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Row')
for row in rows:
    tmp = []
    cells = row.findall('.//{urn:schemas-microsoft-com:office:spreadsheet}Data')
    for cell in cells:
        tmp.append(cell.text)
    L.append(tmp)

df = pd.DataFrame(L[1:], columns=L[0])

df.replace(["0",None], np.nan, inplace=True)

onlyfiles = [f for f in listdir("C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/ReportShopSmartSkill_Accessi") if isfile(join("C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/ReportShopSmartSkill_Accessi", f))]

lastFile = onlyfiles[-1]
lastDatetime = pd.to_datetime(lastFile.split("_")[2].split(".")[0], format='%Y-%m-%d') + pd.Timedelta(1, unit="Day")

df["Data"] = pd.to_datetime(df["Data"], format="%d/%m/%Y %H:%M:%S")
df["Data mod."] = pd.to_datetime(df["Data mod."], format="%d/%m/%Y %H:%M:%S")

df[["Totale","Rinnovo","PRODOTTITOTALI"]] = df[["Totale","Rinnovo","PRODOTTITOTALI"]].apply(lambda x: x.str.replace(',','.'))
df[["Totale","Rinnovo","PRODOTTITOTALI"]] = df[["Totale","Rinnovo","PRODOTTITOTALI"]].fillna(0)
df[["Totale","Rinnovo","PRODOTTITOTALI"]] = df[["Totale","Rinnovo","PRODOTTITOTALI"]].astype(float)

df["Data"] = pd.to_datetime(df["Data"].dt.date)

df = df.loc[(df["EMAIL"] != "ntt.eshop@mailinator.com") & (df["Data"] >= lastDatetime) & (df["Data"].dt.date < dt.date.today()), :]


### ReportShopSmartSkill_Accessi

ReportShopSmartSkill_Accessi = df.loc[df["IDUSER"].isna(), ["#Ordine", "Data"]].groupby(df["Data"]).agg({"#Ordine":"nunique"}).reset_index().rename(columns={"#Ordine":"N. accessi"})
aggiornamento = str(ReportShopSmartSkill_Accessi["Data"].max())[0:10]
ReportShopSmartSkill_Accessi.to_excel(f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/ReportShopSmartSkill_Accessi/ReportShopSmartSkill_Accessi_{aggiornamento}.xlsx", index=False)



### ReportShopSmartSkill

ReportShopSmartSkill = df.loc[df["IDUSER"].notna(), ["Data","#Ordine","IDUSER","EMAIL","Rag.soc","P.IVA","Attivazione", "Prodotti","PRODOTTITOTALI","Totale","Rinnovo"]]
ReportShopSmartSkill.columns = ["Data", "ID Ordine", "ID User", "Email", "Rag.Soc", "P.IVA", "Attivazione", "Prodotti", "N. prodotti", "Valore a carrello", "Valore di rinnovo"]
ReportShopSmartSkill.to_excel(f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/ReportShopSmartSkill/ReportShopSmartSkill_{aggiornamento}.xlsx", index=False)


### ReportShopSmartSkill_Dettaglio

ReportShopSmartSkill_Dettaglio = (df.loc[df["Prodotti"].notna(),["Data", "Attivazione", "Prodotti"]].set_index(["Data", "Attivazione"]).apply(lambda x: x.str.split(';').explode()).reset_index())
ReportShopSmartSkill_Dettaglio["Prodotti"] = ReportShopSmartSkill_Dettaglio["Prodotti"].str.strip()
ReportShopSmartSkill_Dettaglio = ReportShopSmartSkill_Dettaglio.groupby([ReportShopSmartSkill_Dettaglio["Data"], "Attivazione","Prodotti"]).agg({"Prodotti":"count"}).rename(columns={"Prodotti":"N. prodotti"}).reset_index()
ReportShopSmartSkill_Dettaglio.to_excel(f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data source/ReportShopSmartSkill_Dettaglio/ReportShopSmartSkill_Dettaglio_{aggiornamento}.xlsx", index=False)


### Webdesk

onlyfiles = [f for f in sorted(listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data export\StatsLog-SMARTSKILL")) if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data export\StatsLog-SMARTSKILL", f))]

lastFile = onlyfiles[-1]
data = lastFile.split("-")[2].split(".")[0]


zip_file =f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data export/StatsLog-SMARTSKILL/StatsLog-SMARTSKILL-{data}.zip"
directory_to_extract_to = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data source\StatsLogSMARTSKILL"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")


zip_file =f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/E-commerce/Data export/StatsLog-SMARTSKILL/StatsLog-SMARTSKILLDETT-{data}.zip"
directory_to_extract_to = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\E-commerce\Data source\StatsLogSMARTSKILLDETT"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")
