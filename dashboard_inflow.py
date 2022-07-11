import pandas as pd
import datetime as dt
import numpy as np
import calendar
import zipfile

zip_file =r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export\ZQ_MC_BUDANAL_IMP_ALL.zip"
directory_to_extract_to = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export"
try:
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)
except:
    print("Invalid file")

bw_file = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Data export\ZQ_MC_BUDANAL_IMP_ALL_00000.xls"
df = pd.DataFrame(pd.read_html(str(bw_file), encoding = 'utf-8', decimal=",", thousands='.')[3])
df2 = pd.DataFrame(data=df.loc[1:,:])
df2.columns = df.loc[0,:].to_list()
df2["Data Ins. Ordine (monitoraggio)"] = pd.to_datetime(df2["Data Ins. Ordine (monitoraggio)"], format="%d%m%Y")
df2[["Raccolto", "Budget assegnato"]] = df2.loc[:,["Raccolto", "Budget assegnato"]].astype(float)
df2.loc[(df2["Budget assegnato"].notna()) & (df2["Linea Prodotti"] == "Software"), ["Gruppo Prodotti Ricl"]] = "Software"
df2.loc[(df2["Budget assegnato"].notna()) & (df2["Linea Prodotti"] == "Services"), ["Gruppo Prodotti Ricl"]] = "Vendite Services"
df2.rename(columns={'Budget assegnato': 'Budget  assegnato'}, inplace=True)
df2.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\inflow_prodotti\inflow_prodotti_nuove_vendite_2022.xlsx", sheet_name='Foglio1', index=False)

df2_curr_month = df2.loc[df2["Data Ins. Ordine (monitoraggio)"].dt.date >= dt.date(dt.datetime.now().year,dt.datetime.now().month,1),:]
df2_curr_month.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Source\po_giornaliero\inflow_prodotti_nuove_vendite_2022.xlsx", sheet_name='Foglio1', index=False)

gg_lav = pd.DataFrame(data=pd.date_range(start='2021-01-01', end=(dt.date.today()-pd.Timedelta(days=1))), columns=["Giorno"])

holidays=["2021-01-01","2021-01-06","2021-04-05","2021-05-01","2021-06-02","2021-08-15","2021-11-01","2021-12-08","2021-12-25","2021-12-26"]

gg_lav['N. gg lavorativi'] = np.busday_count(
    gg_lav["Giorno"].apply(lambda x: (pd.to_datetime(dt.date(x.year,x.month,1))) ).values.astype('datetime64[D]'),
    (gg_lav['Giorno']+pd.Timedelta(days=1)).values.astype('datetime64[D]'),
    weekmask=[1,1,1,1,1,0,0], holidays=holidays)

gg_lav['N. gg lavorativi nel mese'] = np.busday_count(
    gg_lav["Giorno"].apply(lambda x: (pd.to_datetime(dt.date(x.year,x.month,1))) ).values.astype('datetime64[D]'),
    (gg_lav['Giorno'].apply(lambda d: dt.date(d.year, d.month, calendar.monthrange(d.year, d.month)[-1]))).values.astype('datetime64[D]'),
    weekmask=[1,1,1,1,1,0,0], holidays=holidays)

gg_lav.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Parametriche\Giorni_lavorativi.xlsx", index=False)
