from os import listdir
from os.path import isfile, join
import sys
import win32com.client 

shell = win32com.client.Dispatch("WScript.Shell")


#User directory
user = r"C:\Users\Raffaele.Sportiello"
onedrive_documents = user+r"\OneDrive - Wolters Kluwer\Documents"
pictures = user+r"\Pictures"

#Base dati
base_dati = onedrive_documents+r"\Base dati"

abbonamenti = base_dati+r"\Abbonamenti"
abbonamenti_disdetti = base_dati+r"\Abbonamenti disdetti"
anagrafiche = base_dati+r"\Anagrafiche"
budest = base_dati+r"\Budest"
budget = base_dati+r"\Budget"
dimensioni = base_dati+r"\Dimensioni"
forecast = base_dati+r"\Forecast"
nuovi_clienti = base_dati+r"\Nuovi clienti"
report_delta = base_dati+r"\Report delta"
parametriche = shell.CreateShortCut(base_dati+r"\Parametriche.lnk").Targetpath
salesforce_export = shell.CreateShortCut(base_dati+r"\Salesforce.lnk").Targetpath
zip_folder = base_dati+r"\ZIP"

#Abbonamenti

abb_files = [f for f in listdir(abbonamenti) if isfile(join(abbonamenti, f))]

#Anagrafiche

anag_files = [f for f in listdir(anagrafiche) if isfile(join(anagrafiche, f))]

#Budest
budest_cli_prodotto_anno = budest+r"\Budest cliente prodotto anno"
budest_cli_prodotto_anno_files = [f for f in listdir(budest_cli_prodotto_anno) if isfile(join(budest_cli_prodotto_anno, f))]

budest_prodotto_anno = budest+r"\Budest prodotto anno"
budest_prodotto_anno_files = [f for f in listdir(budest_prodotto_anno) if isfile(join(budest_prodotto_anno, f))]


#Parametriche
parametriche_files = [f for f in listdir(parametriche) if isfile(join(parametriche, f))]

par_file_path={}
for i,elem in enumerate(parametriche_files):
    par_file_path[f'{elem.split(".")[0].lower()}'] = parametriche+"\\"+parametriche_files[i]

#ZIP
zip = base_dati+r"\ZIP"

# Dashboard inflow
dashboard_inflow = onedrive_documents+r"\Dashboard inflow"
dashboard_inflow_canali_prodotti = dashboard_inflow+r"\Dashboard inflow canali e prodotti"

#Dashboard inflow - BPM
bpm_folder = dashboard_inflow_canali_prodotti+r"\BPM"
bpm_folder_export = bpm_folder+r"\Data export"
bpm_export_files = [f for f in listdir(bpm_folder_export) if isfile(join(bpm_folder_export, f))]
bpm_export_file = bpm_folder_export.replace("\\","/")+f"/{bpm_export_files[-1]}"

#Dashboard inflow - Source
dashboard_inflow_source = dashboard_inflow_canali_prodotti+r"\Source"
dashboard_inflow_prodotti = dashboard_inflow_source+r"\inflow_prodotti"
dashboard_inflow_prodotti_files = [f for f in listdir(dashboard_inflow_prodotti) if isfile(join(dashboard_inflow_prodotti, f))]

#Dashboard inflow - Confronto
dashboard_inflow_confronto = dashboard_inflow_canali_prodotti+r"\Confronto"
dashboard_inflow_ieri = dashboard_inflow_confronto+r"\Ieri"
dashboard_inflow_oggi = dashboard_inflow_confronto+r"\Oggi"
dashboard_inflow_po = dashboard_inflow_confronto+r"\Po"

#Salesfoce
salesforce = onedrive_documents+r"\Salesforce"
accounts_sf_file = salesforce_export+r"\IT_all_accounts.csv"
op_w_prod_file = salesforce_export+r"\IT_Op_w_Products.csv"
forecast_cat_w_op_file = salesforce_export+r"\IT_Forecast_Category_w_Opp.csv"