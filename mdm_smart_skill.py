import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join

inflow_prodotti = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti\\"

onlyfiles = [f for f in listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti") if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti", f))]

df = pd.DataFrame()

for file in onlyfiles:
    data = pd.read_excel(inflow_prodotti+file)
    df = df.append(data)

for col in df.select_dtypes(include=[object]).columns:
    df[col] = df[col].astype(str)

#-----------------------------------------------------------------------------------------------------------------------------
parametriche = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche\\"
onlyfiles = [f for f in listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche") if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche", f))]

par = {}
for file in onlyfiles:
    file_name=file.split(".")[0]
    par[f"{file_name}"] = pd.read_excel(parametriche+file)

#-----------------------------------------------------------------------------------------------------------------------------
columns = ["Cliente Merce", "Sales Director", "Canale di VENDITA", "Linea Prodotti", "Codice PRODOTTO", "Numero Ordine", "Numero Fattura", "Data Ins. Ordine (monitoraggio)", "Raccolto"]

df_smart = df.loc[:, columns].reset_index(drop=True)

df_smart["Sales Director"] = df_smart["Sales Director"].str.lower()
df_smart["Codice PRODOTTO"] = df_smart["Codice PRODOTTO"].astype(str)

esc_fatture = par["Esclusione_fatture"]
esc_fatture["Numero Fattura"] = esc_fatture["Numero Fattura"].astype(str)

esc_ordini = par["Esclusione_ordini"]
esc_ordini["Numero Ordine"] = esc_ordini["Numero Ordine"].astype(str)

sales_director = par["Sales_director"]
sales_director["Sales Director"] = sales_director.loc[:,"Sales Director"].str.lower()
sales_director = sales_director[["Sales Director", "Rete", "Rete3", "Area territoriale"]]

prodotti = par["Classificazione_codice_prodotto"]
prodotti["Codice PRODOTTO"] = prodotti["Codice PRODOTTO"].astype(str)

li_prodotti = par["Linea_prodotti"]

clienti_smart = df_smart.merge(esc_fatture, on="Numero Fattura", how="left").merge(esc_ordini, on="Numero Ordine", how="left").merge(li_prodotti, on="Linea Prodotti", how="left").merge(sales_director, on="Sales Director", how="left").merge(prodotti, on="Codice PRODOTTO", how="left")

esclusioni = (clienti_smart["Escludi fatture"].isna())&(clienti_smart["Escludi ordine"].isna())&(clienti_smart["Sub solution"] == "Smart - Skills")

clienti_smart = clienti_smart.loc[esclusioni, ['Data Ins. Ordine (monitoraggio)', "Cliente Merce", "Numero Ordine", "Rete3", "Canale di VENDITA",'Linea', 'Raccolto']].reset_index(drop=True)

clienti_smart.rename(columns={"Data Ins. Ordine (monitoraggio)":'Giorno'}, inplace=True)
#-----------------------------------------------------------------------------------------------------------------------------

arr = clienti_smart["Giorno"].astype('datetime64[M]').sort_values().unique()
d_rete = {}
d_canale = {}

for i,elem in enumerate(arr):
    temp = clienti_smart.loc[(clienti_smart["Giorno"].astype('datetime64[M]') >= "2021-04-01") &
                             (clienti_smart["Giorno"].astype('datetime64[M]') <= elem) &
                             (clienti_smart["Canale di VENDITA"] != "E-Commerce"), :].groupby(["Rete3"]).agg({"Cliente Merce":"nunique"}).reset_index()
    d_rete[f"d_{i}"] = temp["Cliente Merce"].to_numpy()

    temp = clienti_smart.loc[(clienti_smart["Giorno"].astype('datetime64[M]') >= "2021-04-01") &
                             (clienti_smart["Giorno"].astype('datetime64[M]') <= elem) &
                             (clienti_smart["Canale di VENDITA"] == "E-Commerce"), :].groupby(["Canale di VENDITA"]).agg({"Cliente Merce":"nunique"}).reset_index()
    d_canale[f"d_{i}"] = temp["Cliente Merce"].to_numpy()

df_rete = pd.DataFrame.from_dict(data= d_rete.items())
df_rete = pd.DataFrame(df_rete[1].tolist())
df_rete = df_rete.pivot_table(columns=df_rete.index)

df_canale = pd.DataFrame.from_dict(data= d_canale.items())
df_canale = pd.DataFrame(df_canale[1].tolist())
df_canale = df_canale.pivot_table(columns=df_canale.index)

clienti_unici = pd.concat([df_rete, df_canale], axis=0)
clienti_unici.columns = arr
clienti_unici.index = ["Diretta", "Indiretta", "E-commerce"]
clienti_unici = clienti_unici.transpose()
clienti_unici["Traditional channels"] = clienti_unici["Diretta"] + clienti_unici["Indiretta"]
clienti_unici = clienti_unici.loc["2022-01-01":, ["Traditional channels", "E-commerce"]]
#-----------------------------------------------------------------------------------------------------------------------------

inflow_progr_Anna = clienti_smart.loc[(clienti_smart["Canale di VENDITA"] != "E-Commerce") & (clienti_smart["Giorno"] >= "2022-01-01"), :].groupby([clienti_smart["Giorno"].astype("datetime64[M]")])["Raccolto"].sum().reset_index()
inflow_progr_Anna["Anna"] = inflow_progr_Anna["Raccolto"].cumsum()

inflow_progr_Ben = clienti_smart.loc[(clienti_smart["Canale di VENDITA"] == "E-Commerce") & (clienti_smart["Giorno"] >= "2022-01-01"), :].groupby([clienti_smart["Giorno"].astype("datetime64[M]")])["Raccolto"].sum().reset_index()
inflow_progr_Ben["Ben"] = inflow_progr_Ben["Raccolto"].cumsum()

inflow_progr = inflow_progr_Anna.merge(inflow_progr_Ben, how="outer", on="Giorno")
inflow_progr = inflow_progr.loc[:,["Giorno","Anna","Ben"]]
#-----------------------------------------------------------------------------------------------------------------------------

with pd.ExcelWriter(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Presentazioni\MDM\Smart SKILL\clienti_e_inflow_per_slide.xlsx", engine="openpyxl") as writer:
    inflow_progr.to_excel(writer, sheet_name="inflow_progr", index=False)
    clienti_unici.to_excel(writer, sheet_name="clienti_unici", index=True)
