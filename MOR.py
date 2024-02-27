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
