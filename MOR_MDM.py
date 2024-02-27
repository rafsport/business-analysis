import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join
from directories import *

#-----------------------------------------------------------------------------------------------------------------------------
# IT_Op_w_Products
#-----------------------------------------------------------------------------------------------------------------------------
op_w_prod = pd.read_csv(op_w_prod_file, encoding='latin-1', sep=";")

for col in op_w_prod.columns:
    if "Date" in col:
        op_w_prod[col] = op_w_prod[col].astype(str)
        op_w_prod[col] = pd.to_datetime(op_w_prod[col], format="%d/%m/%Y")
    else:
        pass

columns = [c for c in op_w_prod.columns if not "Currency" in c]
op_w_prod = op_w_prod.loc[:,columns]

num = []

for col in op_w_prod.columns:
    if "Amount" in col:
        num.append(col)
    else:
        pass
    
op_w_prod[num] = op_w_prod[num].apply(lambda x: x.str.replace(",",".", regex=False)).astype(float)

op_w_prod["weighted amount"] = op_w_prod["weighted amount"].apply(lambda x: x.replace(",",".")).astype(float)
op_w_prod["Total Value"] = op_w_prod["Total Value"].apply(lambda x: x.replace(",",".")).astype(float)

op_w_prod["Total Price"] = op_w_prod["Total Price"].astype(str).apply(lambda x: x.replace(",",".")).astype(float)

#-----------------------------------------------------------------------------------------------------------------------------
# Success per MOR
#-----------------------------------------------------------------------------------------------------------------------------
now = dt.datetime.now()
prev_month = now.month-1
curr_month = now.month
current_year = now.year

columns = ["Opportunity ID","Opportunity Name", "Product Name", "Total Price", "Account Name",  "Owner Role", "Close Date","Stage", "Incremental Amount"]

filt = (op_w_prod["Stage"].str.contains("Closed Won")) & (op_w_prod["Close Date"].dt.month==(prev_month)) & (op_w_prod["Close Date"].dt.year==(current_year))

closed_deals_con_prod = op_w_prod.loc[filt,:].groupby(["Opportunity ID","Opportunity Name","Account Name",  "Owner Role", "Close Date", "Product Name"]).agg({"Total Price":"sum","Incremental Amount":"max"}).sort_values(by=["Incremental Amount"], ascending=False)

closed_deals_con_prod.style.format("{:,.0f}€").background_gradient().to_excel(onedrive_documents + r"\Presentazioni\MDM\Success per MOR.xlsx")



filt = (op_w_prod["Stage"].str.contains("Closed Won")) & (op_w_prod["Close Date"].dt.month==(curr_month)) & (op_w_prod["Close Date"].dt.year==(current_year))

closed_deals_con_prod = op_w_prod.loc[filt,:].groupby(["Opportunity ID","Opportunity Name","Account Name",  "Owner Role", "Close Date", "Product Name"]).agg({"Total Price":"sum","Incremental Amount":"max"}).sort_values(by=["Incremental Amount"], ascending=False)

closed_deals_con_prod.style.format("{:,.0f}€").background_gradient().to_excel(onedrive_documents + r"\Presentazioni\MDM\Success per MOR - current month.xlsx")

#-----------------------------------------------------------------------------------------------------------------------------
# Smart SKILL # 
#-----------------------------------------------------------------------------------------------------------------------------
dfs = []
for file in dashboard_inflow_prodotti_files:
    data = pd.read_excel(dashboard_inflow_prodotti+"\\"+file)
    dfs.append(data)
inflow = pd.concat(dfs, axis=0, ignore_index=True)

for col in inflow.select_dtypes(include=[object]).columns:
    inflow[col] = inflow[col].astype(str)

#-----------------------------------------------------------------------------------------------------------------------------
par = {}
for file in parametriche_files:
    file_name=file.split(".")[0]
    par[f"{file_name}"] = pd.read_excel(parametriche+"\\"+file)

#-----------------------------------------------------------------------------------------------------------------------------
columns = ["Cliente Merce", "Sales Director", "Canale di VENDITA", "Linea Prodotti", "Codice PRODOTTO", "Numero Ordine", "Numero Fattura", "Data Ins. Ordine (monitoraggio)", "Raccolto"]

df_smart = inflow.loc[:, columns].reset_index(drop=True)

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

arr = clienti_smart.loc[clienti_smart["Giorno"].dt.year == current_year, "Giorno"].astype('datetime64[M]').sort_values().unique()
d_rete = {}
d_canale = {}

for i,elem in enumerate(arr):
    temp = clienti_smart.loc[(clienti_smart["Giorno"].astype('datetime64[M]').dt.date >= dt.date(current_year,1,1)) &
                             (clienti_smart["Giorno"].astype('datetime64[M]') <= elem) &
                             (clienti_smart["Canale di VENDITA"] != "E-Commerce"), :].groupby(["Rete3"]).agg({"Cliente Merce":"nunique"}).reset_index()
    d_rete[f"d_{i}"] = temp["Cliente Merce"].to_numpy()

    temp = clienti_smart.loc[(clienti_smart["Giorno"].astype('datetime64[M]').dt.date >= dt.date(current_year,1,1)) &
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
clienti_unici = clienti_unici.loc[dt.date(current_year,1,1):, ["Traditional channels", "E-commerce"]]
#-----------------------------------------------------------------------------------------------------------------------------

inflow_progr_Anna = clienti_smart.loc[(clienti_smart["Canale di VENDITA"] != "E-Commerce") & 
                                      (clienti_smart["Giorno"].dt.date >= dt.date(current_year,1,1)), 
                                      :].groupby([clienti_smart["Giorno"].astype("datetime64[M]")])["Raccolto"].sum().reset_index()
inflow_progr_Anna["Anna"] = inflow_progr_Anna["Raccolto"].cumsum()

inflow_progr_Ben = clienti_smart.loc[(clienti_smart["Canale di VENDITA"] == "E-Commerce") &
                                    (clienti_smart["Giorno"].dt.date >= dt.date(current_year,1,1)), 
                                    :].groupby([clienti_smart["Giorno"].astype("datetime64[M]")])["Raccolto"].sum().reset_index()
inflow_progr_Ben["Ben"] = inflow_progr_Ben["Raccolto"].cumsum()

inflow_progr = inflow_progr_Anna.merge(inflow_progr_Ben, how="outer", on="Giorno")
inflow_progr = inflow_progr.loc[:,["Giorno","Anna","Ben"]]
#-----------------------------------------------------------------------------------------------------------------------------

with pd.ExcelWriter(onedrive_documents + r"\Presentazioni\MDM\Smart SKILL\clienti_e_inflow_per_slide.xlsx", engine="openpyxl") as writer:
    inflow_progr.to_excel(writer, sheet_name="inflow_progr", index=False)
    clienti_unici.to_excel(writer, sheet_name="clienti_unici", index=True)

#-----------------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------
