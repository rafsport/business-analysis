#############################################################################################################
# Elaborazione

import pandas as pd
import os
import numpy as np
import datetime as dt
from os import listdir
from os.path import isfile, join

onlyfiles = [f for f in listdir(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Inserimenti manuali\Data export") if isfile(join(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Inserimenti manuali\Data export", f))]

lastFile = onlyfiles[-1]

ins = pd.read_excel(f"C:/Users/Raffaele.Sportiello/OneDrive - Wolters Kluwer/Documents/Dashboard inflow/Dashboard inflow canali e prodotti/Inserimenti manuali/Data export/{lastFile}")

ins["Data di completamento"] = ins["Ora di completamento"].dt.date

ins["Qual è il tipo di modifica manuale richiesta?"] = np.where(ins["Qual è il tipo di modifica manuale richiesta?"] == "Cambio dell'intestazione di un ordine", ins["Qual è il tipo di cambio che si vuole operare?"], ins["Qual è il tipo di modifica manuale richiesta?"])

ins["Sales Director"] = ins.loc[:,ins.columns.str.contains("Su quale sales director")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Agente dell'ORDINE"] = ins.loc[:,ins.columns.str.contains("Su quale codice agente")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Codice PRODOTTO"] = ins.loc[:,ins.columns.str.contains("Su quali codici prodotto")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Linea Prodotto"] = ins.loc[:,ins.columns.str.contains("Su quale linea prodotto")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Raccolto"] = ins.loc[:,ins.columns.str.contains("valore dell'ordine")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Data"] = ins.loc[:,ins.columns.str.contains("quale data")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Numero Ordine"] = ins.loc[:,ins.columns.str.contains("numero dell'ordine")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Numero Fattura"] = ins.loc[:,ins.columns.str.contains("numero della fattura")].fillna(axis=1, method="ffill").iloc[:,-1]
ins["Testo Richiesta"] = ins.loc[:,ins.columns.str.contains("testo della richiesta")].fillna(axis=1, method="ffill").iloc[:,-1]

ins = ins.loc[:,["ID", "Nome", "Data di completamento", "Qual è il tipo di modifica manuale richiesta?",  "Numero Ordine", "Numero Fattura", "Sales Director", "Agente dell'ORDINE", "Codice PRODOTTO", "Linea Prodotto", "Data", "Testo Richiesta", "Raccolto"]]

ins["Linea Prodotto"] = ins["Linea Prodotto"].str.replace("10,4k Software; 6,1k Servizi","Software,Servizi")

ins = (ins.set_index(["ID", "Nome", "Data di completamento", "Testo Richiesta", "Qual è il tipo di modifica manuale richiesta?", "Numero Ordine", "Numero Fattura", "Sales Director","Agente dell'ORDINE", "Data"]).apply(lambda x: x.str.split(';|,').explode()).reset_index())

ins["Raccolto"] = ins["Raccolto"].astype(float)

ins.to_excel(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Inserimenti manuali\Inserimenti_manuali.xlsx")


import win32com.client as client
#############################################################################################################
# Email
#
inserimenti_manuali = r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Dashboard inflow\Dashboard inflow canali e prodotti\Inserimenti manuali\Inserimenti_manuali.xlsx"

html_body = """
    <div>
          <p>Ciao Loredana, Rocco,<br><br>
            in allegato il file aggiornato.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Loredana.Montagna@wolterskluwer.com;Rocco.Tarquini@wolterskluwer.com"
message.Subject = 'Inserimenti manuali'
message.HTMLBody = html_body
message.Attachments.Add(Source=inserimenti_manuali)

message.Display()
