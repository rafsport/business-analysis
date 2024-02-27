import pandas as pd
import datetime as dt
import numpy as np
import calendar
import zipfile
import sys

from os import listdir
from os.path import isfile, join
from os.path import getmtime, getctime
import shutil

from bs4 import BeautifulSoup
import win32com.client
import plotly.graph_objects as go
import plotly.offline as pyo
from directories import *

zip_file =dashboard_inflow_canali_prodotti+r"\Data export\ZQ_MC_BUDANAL_IMP_ALL.zip"
directory_to_extract_to = dashboard_inflow_canali_prodotti+r"\Data export"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")

bw_file = dashboard_inflow_canali_prodotti+r"\Data export\ZQ_MC_BUDANAL_IMP_ALL_00000.xls"
df = pd.DataFrame(pd.read_html(str(bw_file), encoding = 'utf-8', decimal=",", thousands='.')[3])
df2 = pd.DataFrame(data=df.loc[1:,:])
df2.columns = df.loc[0,:].to_list()
df2["Data Ins. Ordine (monitoraggio)"] = pd.to_datetime(df2["Data Ins. Ordine (monitoraggio)"], format="%d%m%Y")
df2[["Raccolto", "Budget assegnato"]] = df2.loc[:,["Raccolto", "Budget assegnato"]].astype(float)
df2.loc[(df2["Budget assegnato"].notna()) & (df2["Linea Prodotti"] == "Software"), ["Gruppo Prodotti Ricl"]] = "Software"
df2.loc[(df2["Budget assegnato"].notna()) & (df2["Linea Prodotti"] == "Service on-line"), ["Gruppo Prodotti Ricl"]] = "Software"
df2.loc[(df2["Budget assegnato"].notna()) & (df2["Linea Prodotti"] == "Services"), ["Gruppo Prodotti Ricl"]] = "Vendite Services"
df2.rename(columns={'Budget assegnato': 'Budget  assegnato'}, inplace=True)


# Elenco ordini inseriti nel giorno

df2_old = pd.read_excel(dashboard_inflow_prodotti+r"\inflow_prodotti_nuove_vendite_2023.xlsx")
inflow_giorno = df2.merge(df2_old, how="outer", indicator=True)
inflow_giorno = inflow_giorno.loc[inflow_giorno["_merge"] == "left_only",:]
inflow_giorno.loc[:,['Codice PRODOTTO','Gruppo Prodotti Ricl', "Agente dell'ORDINE",'Sales Director', 'Canale di VENDITA', 'Cliente Merce','Cliente FATTURA', 'Linea Prodotti', 'Numero Ordine', 'Origine Ordine','Numero Fattura', 'Data Ins. Ordine (monitoraggio)', 'Raccolto']].to_excel(dashboard_inflow_source+r"\Ordini inseriti nel giorno.xlsx", index=False)

df2.to_excel(dashboard_inflow_prodotti+r"\inflow_prodotti_nuove_vendite_2024.xlsx", sheet_name='Foglio1', index=False)


# Elenco ordini SAP per confronto con BPM

df2.groupby(["Numero Ordine","Data Ins. Ordine (monitoraggio)"], dropna=False).agg({"Raccolto":'sum'}).reset_index().to_excel(bpm_folder+r"\Ordini_SAP.xlsx", sheet_name="Foglio1", index=False)


# Inflow mese corrente per PO Giornaliero

df2_curr_month = df2.loc[df2["Data Ins. Ordine (monitoraggio)"].dt.date >= dt.date(dt.datetime.now().year,dt.datetime.now().month,1),:]
df2_curr_month.to_excel(dashboard_inflow_source+f"/po_giornaliero/inflow_prodotti_nuove_vendite_2024_mese_corrente.xlsx", sheet_name='Foglio1', index=False)

gg_lav = pd.DataFrame(data=pd.date_range(start='2021-01-01', end=(dt.date.today()-pd.Timedelta(days=1))), columns=["Giorno"])


# Calcolo dei giorni lavorativi per avanzamento

holidays=["2023-01-01","2023-01-06","2023-04-10","2023-05-01","2023-06-02","2023-08-15","2023-11-01","2023-12-08","2023-12-25","2023-12-26"]

gg_lav['N. gg lavorativi'] = np.busday_count(
    gg_lav["Giorno"].apply(lambda x: (pd.to_datetime(dt.date(x.year,x.month,1))) ).values.astype('datetime64[D]'),
    (gg_lav['Giorno']+pd.Timedelta(days=1)).values.astype('datetime64[D]'),
    weekmask=[1,1,1,1,1,0,0], holidays=holidays)

gg_lav['N. gg lavorativi nel mese'] = np.busday_count(
    gg_lav["Giorno"].apply(lambda x: (pd.to_datetime(dt.date(x.year,x.month,1))) ).values.astype('datetime64[D]'),
    (gg_lav['Giorno'].apply(lambda d: dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[-1]))).values.astype('datetime64[D]'),
    weekmask=[1,1,1,1,1,0,0], holidays=holidays)

gg_lav.to_excel(parametriche+r"\Giorni_lavorativi.xlsx", index=False)




# CREAZIONE GRAFICO AVANZAMENTO INFLOW VS. FORECAST

# Calcolo dell'obiettivo
obiettivo = gg_lav.assign(Obiettivo = round(gg_lav["N. gg lavorativi"]/gg_lav["N. gg lavorativi nel mese"],2)).loc[:,"Obiettivo"].tail(1).values[0]
inflow = pd.concat([df2, 
                    pd.read_excel(dashboard_inflow_prodotti+"\\"+dashboard_inflow_prodotti_files[0])], axis=0, ignore_index=True)
forecast = pd.read_excel(dashboard_inflow_canali_prodotti+r"\Source\forecast_sales_director\forecast_sales_director_2023.xlsx")

par = {}
for file in parametriche_files:
    file_name=file.split(".")[0]
    par[f"{file_name}"] = pd.read_excel(parametriche+"\\"+file)
sales_director = pd.DataFrame(par["Sales_director"]).loc[:,["Sales Director", "Area territoriale", "Rete"]]

li_prodotti = pd.DataFrame(par["Linea_prodotti"])
esc_fatture = pd.DataFrame(par["Esclusione_fatture"])
esc_ordini = pd.DataFrame(par["Esclusione_ordini"])

inflow.rename(columns={"Data Ins. Ordine (monitoraggio)":"Giorno"}, inplace=True)
forecast.rename(columns={"Sales director":"Sales Director"}, inplace=True)


# Creazione del Dataframe in merge con il Forecast

inflow["Giorno"] = pd.to_datetime(inflow["Giorno"], format='%Y-%m-%d')
current_month = pd.to_datetime(dt.date.today().replace(day=1)).to_period('M')

df_merged = (inflow.loc[(inflow["Giorno"].dt.to_period('M') == current_month) &
                    (~inflow["Numero Ordine"].isin(esc_ordini["Numero Ordine"])) & 
                    (~inflow["Numero Fattura"].isin(esc_fatture["Numero Fattura"])) & 
                    (inflow["Canale di VENDITA"] != "E-Commerce"),
                    ["Giorno", "Canale di VENDITA", "Numero Fattura", "Numero Ordine", "Sales Director", "Raccolto"]]\
            .merge(sales_director, how="left", on="Sales Director"))

df_merged["Giorno"] = pd.to_datetime(df_merged["Giorno"], format='%Y-%m-%d').dt.to_period('M')
forecast["Giorno"] = pd.to_datetime(forecast["Giorno"], format='%Y-%m-%d').dt.to_period('M')
    
df_avanz = df_merged.groupby(["Giorno", "Rete", "Area territoriale"], dropna=False).agg({"Raccolto":"sum"}).reset_index()\
.merge((forecast.merge(sales_director, how="left", on="Sales Director"))\
.groupby(["Giorno", "Rete", "Area territoriale"], dropna=False).agg({"Forecast":"sum"}).reset_index())
    
df_avanz = pd.concat([df_avanz, df_avanz[["Raccolto","Forecast"]].agg("sum").rename("Totale").to_frame().T], axis=0)\
.assign(Avanz = lambda x: np.where(x["Forecast"] > 0, x["Raccolto"]/x["Forecast"], 0))
    
df_avanz.loc[df_avanz.index == "Totale", ["Giorno", "Rete", "Area territoriale"]] = [pd.to_datetime(dt.date.today().replace(day=1)).to_period('M'), "Totale", ""]

df = df_avanz.sort_values(by=["Rete", "Area territoriale"], ascending=False).copy()


# Creazione del grafico

fig = go.Figure()


fig.add_trace(go.Bar(
    y=df.loc[df["Rete"] =="Totale","Rete"] ,
    x=df.loc[df["Rete"] =="Totale", "Avanz"],
    orientation='h',
    marker=dict(color='#003566'),
    marker_line=dict(width=3, color='#219ebc'),
    name='Avanzamento',
    texttemplate='%{x:.0%}',
    textposition='inside'
))

# Aggiunta della barra orizzontale
fig.add_trace(go.Bar(
    y=df.loc[df["Rete"] !="Totale","Rete"] + " - " + df.loc[df["Rete"] !="Totale", "Area territoriale"],
    x=df.loc[df["Rete"] !="Totale", "Avanz"],
    orientation='h',
    marker=dict(color='#003566'),
    name='Avanzamento',
    texttemplate='%{x:.0%}',
    textposition='inside'
))



# Aggiunta della linea verticale dell'obiettivo
fig.add_shape(
    type="line",
    x0=obiettivo,
    y0=-0.5,
    x1=obiettivo,
    y1=len(df)-0.5,
    line=dict(
        color="red",
        width=3,
        dash="dashdot"
    )
)

# Impostazioni layout
fig.update_layout(
    title= dict(text="Avanzamento inflow vs. forecast", font_size=24),
    #xaxis_title_text="Avanzamento",
    yaxis_title_text="Rete - Area territoriale",
    yaxis=dict(automargin=True),
    xaxis=dict(tickformat=',.0%'),
    xaxis_range=[0,1],
    #category_orders={'Rete': category_order},
    height=400,
    margin=dict(l=200, r=50, t=100, b=100),
    template='plotly_white',
    showlegend=False
)

# Mostra il grafico
fig.write_image(pictures+r'\Avanzamento_inflow_vs_forecast.png',width=1000, height=500 )
pyo.plot(fig, filename=dashboard_inflow_canali_prodotti+f'/Avanzamento_inflow_vs_forecast.html')



# INDIVIDUAZIONE NUOVI CODICI PRODOTTO

zip_file = dashboard_inflow+r"\Lista prodotti\ZQ_MC_BUDANAL_IMP_ALL.zip"
directory_to_extract_to = dashboard_inflow + r"\Lista prodotti"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")


cod_prod = pd.read_excel(parametriche+r"\Classificazione_codice_prodotto.xlsx")

nuovi = df2.loc[~df2["Codice PRODOTTO"].isin(cod_prod["Codice PRODOTTO"]),:].drop_duplicates()


if nuovi.empty == True:
    raise RuntimeError("Non ci sono nuovi codici prodotto")
    


with open(dashboard_inflow + r"\Lista prodotti\ZQ_MC_BUDANAL_IMP_ALL_00000.xls", 'r') as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

# Recupera la tabella di interesse
table_html = str(soup.find_all('table')[8])

# Correggi il valore di colspan nella cella interessata
table_html = table_html.replace('colspan="3D2"', 'colspan="2"')

# Leggi la tabella HTML corretta con Pandas
table = pd.read_html(table_html)[0].iloc[1:,0:3]

table.columns = ["Codice PRODOTTO", "Descrizione prodotto", "Prodotto (SW)"]


nuovi = nuovi.merge(table, how="left", on="Codice PRODOTTO").reindex(columns=cod_prod.columns).fillna("")


html_body = "<div><p>Ciao Stefania,<br><br>ho trovato questo nuovo codice:<br><br>" + nuovi.to_html(index=False) + "</div><br><br>Grazie,<br>Raffaele<br><br></p></div>"

outlook = win32com.client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "stefania.cremonesi@wolterskluwer.com"
message.CC = "lorella.vigato@wolterskluwer.com;martina.tisa@wolterskluwer.com;Elenamaria.Regazzoni@wolterskluwer.com;dimitri.falzone@wolterskluwer.com"
message.Subject = 'Classificazione nuovi codici prodotto'
message.HTMLBody = html_body

message.Display()