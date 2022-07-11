import pandas as pd
import numpy as np
import datetime as dt
import calendar
import plotly.graph_objects as go
import plotly.offline as pyo
from plotly.subplots import make_subplots
import plotly.express as px

forecast = pd.read_csv(r"C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Salesforce\Data export\IT_Forecast_and_Opportunities.csv", encoding='latin-1', sep=";")

for col in forecast.columns:
    if "Date" in col:
        forecast[col] = forecast[col].astype(str)
        forecast[col] = pd.to_datetime(forecast[col], format="%d/%m/%Y")
    else:
        pass


columns = [c for c in forecast.columns if not "Currency" in c]
forecast = forecast.loc[:,columns]

num = []

for col in forecast.columns:
    if "Amount" in col:
        num.append(col)
    else:
        pass

forecast[num] = forecast[num].apply(lambda x: x.str.replace(",",".", regex=False)).astype(float)

role = ["IT Large Accounts-Enterprise Specialists", "IT Central & South Manager", "IT North West Manager", "IT North East Manager"]
filt = (forecast["Forecasting Type: API Name"] == "OpportunityLineItemRevenue") & \
(forecast["Forecasting Item Category"].isin(["Open Pipeline","Commit Forecast"])) & \
(forecast["Owner: Role: Name"].isin(role) & \
 (forecast["End Date"].dt.date >= dt.datetime.now().date()) & \
 (forecast["End Date"].dt.date < dt.date((dt.datetime.now().date() + dt.timedelta(days=90)).year,
                                        (dt.datetime.now().date() + dt.timedelta(days=90)).month,
                                        calendar.monthrange((dt.datetime.now().date() + dt.timedelta(days=90)).year,
                                                            (dt.datetime.now().date() + dt.timedelta(days=90)).month)[1])))


columns = ["End Date", "Product Family", "Forecasting Item Category", "Forecast Amount", "Owner: Full Name"]

commit_forecast = forecast.loc[filt,columns]

commit_forecast.loc[(commit_forecast["Forecasting Item Category"] == "Commit Forecast")\
                    &(commit_forecast["End Date"] == commit_forecast["End Date"].min()), "Goal"] = commit_forecast["Forecast Amount"] * 3

commit_forecast.loc[(commit_forecast["Forecasting Item Category"] == "Commit Forecast")&\
                    (commit_forecast["End Date"] != commit_forecast["End Date"].min())&\
                    (commit_forecast["End Date"] != commit_forecast["End Date"].max()),
                    "Goal"] = commit_forecast["Forecast Amount"] * 2

commit_forecast.loc[(commit_forecast["Forecasting Item Category"] == "Commit Forecast")&\
                    (commit_forecast["End Date"] == commit_forecast["End Date"].max()), "Goal"] = commit_forecast["Forecast Amount"] * 0.8


pipe = commit_forecast.groupby([commit_forecast["End Date"], "Product Family", "Forecasting Item Category","Owner: Full Name"], dropna=False).agg({"Forecast Amount":"sum","Goal":"sum"}).reset_index()

pipe["End Date"] = pipe["End Date"].dt.strftime('%B')


goal = {}
for name in pipe["Owner: Full Name"].unique():
    for i,month in enumerate(pipe["End Date"].unique()):
        goal[f"{name} {i}"] = [pipe.loc[(pipe["Forecasting Item Category"] == "Commit Forecast")&(pipe["Owner: Full Name"] == name)&(pipe["End Date"] == month), "Goal"].sum(),
                               pipe.loc[(pipe["Forecasting Item Category"] == "Open Pipeline")&(pipe["Owner: Full Name"] == name)&(pipe["End Date"] == month), "Forecast Amount"].sum() -\
                               pipe.loc[(pipe["Forecasting Item Category"] == "Commit Forecast")&(pipe["Owner: Full Name"] == name)&(pipe["End Date"] == month), "Goal"].sum()]


af_g_0 = goal["Andrea Ferrara 0"][0]
af_d_0 = goal["Andrea Ferrara 0"][1]
af_g_1 = goal["Andrea Ferrara 1"][0]
af_d_1 = goal["Andrea Ferrara 1"][1]
af_g_2 = goal["Andrea Ferrara 2"][0]
af_d_2 = goal["Andrea Ferrara 2"][1]

al_g_0 = goal["Aureliano Leone 0"][0]
al_d_0 = goal["Aureliano Leone 0"][1]
al_g_1 = goal["Aureliano Leone 1"][0]
al_d_1 = goal["Aureliano Leone 1"][1]
al_g_2 = goal["Aureliano Leone 2"][0]
al_d_2 = goal["Aureliano Leone 2"][1]

cf_g_0 = goal["Camilla Fabris 0"][0]
cf_d_0 = goal["Camilla Fabris 0"][1]
cf_g_1 = goal["Camilla Fabris 1"][0]
cf_d_1 = goal["Camilla Fabris 1"][1]
cf_g_2 = goal["Camilla Fabris 2"][0]
cf_d_2 = goal["Camilla Fabris 2"][1]

cm_g_0 = goal["Cristiano Mozzolani 0"][0]
cm_d_0 = goal["Cristiano Mozzolani 0"][1]
cm_g_1 = goal["Cristiano Mozzolani 1"][0]
cm_d_1 = goal["Cristiano Mozzolani 1"][1]
cm_g_2 = goal["Cristiano Mozzolani 2"][0]
cm_d_2 = goal["Cristiano Mozzolani 2"][1]


########################################################################################
fig = px.bar(pipe, x="Forecasting Item Category",
             y="Forecast Amount",
             facet_row="End Date",
             facet_col="Owner: Full Name",
             facet_col_wrap=10,
             color="Product Family",
             color_discrete_map={
                "Subscription": "#0600C2",
                "Non-Subscription": "#00C288"},
             category_orders={"Forecasting Item Category": ["Open Pipeline", "Commit Forecast"],
                             "Product Family":["Subscription", "Non-Subscription"]})


fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))


fig.update_layout(dict(
    title=dict(text="Rapporto Pipeline vs. Forecast",
               pad=dict(b=500),
               font=dict(size=24)),
    plot_bgcolor="#DADADA")
                 )
#Ferrara 3° mese
fig.add_hline(y=af_g_2, line_dash="dot", row=1, col=1,
              annotation_text=f"Pipe goal: {af_g_2:,.0f}, Δ: {af_d_2:,.0f} ",
              annotation_position="top right")
#Leone 3° mese
fig.add_hline(y=al_g_2, line_dash="dot", row=1, col=2,
              annotation_text=f"Pipe goal: {al_g_2:,.0f}, Δ: {al_d_2:,.0f} ",
              annotation_position="top right")
#Fabris 3° mese
fig.add_hline(y=cf_g_2, line_dash="dot", row=1, col=3,
              annotation_text=f"Pipe goal: {cf_g_2:,.0f}, Δ: {cf_d_2:,.0f} ",
              annotation_position="top right")
#Mozzolani 3° mese
fig.add_hline(y=cm_g_2, line_dash="dot", row=1, col=4,
              annotation_text=f"Pipe goal: {cm_g_2:,.0f}, Δ: {cm_d_2:,.0f} ",
              annotation_position="top right")

#Ferrara 2° mese
fig.add_hline(y=af_g_1, line_dash="dot", row=2, col=1,
              annotation_text=f"Pipe goal: {af_g_1:,.0f}, Δ: {af_d_1:,.0f} ",
              annotation_position="top right")
#Leone 2° mese
fig.add_hline(y=al_g_1, line_dash="dot", row=2, col=2,
              annotation_text=f"Pipe goal: {al_g_1:,.0f}, Δ: {al_d_1:,.0f} ",
              annotation_position="top right")
#Fabris 2° mese
fig.add_hline(y=cf_g_1, line_dash="dot", row=2, col=3,
              annotation_text=f"Pipe goal: {cf_g_1:,.0f}, Δ: {cf_d_1:,.0f} ",
              annotation_position="top right")
#Mozzolani 2° mese
fig.add_hline(y=cm_g_1, line_dash="dot", row=2, col=4,
              annotation_text=f"Pipe goal: {cm_g_1:,.0f}, Δ: {cm_d_1:,.0f} ",
              annotation_position="top right")


#Ferrara 1° mese
fig.add_hline(y=af_g_0, line_dash="dot", row=3, col=1,
              annotation_text=f"Pipe goal: {af_g_0:,.0f}, Δ: {af_d_0:,.0f} ",
              annotation_position="top right",
              line_color="#C28800")
#Leone 3° mese
fig.add_hline(y=al_g_0, line_dash="dot", row=3, col=2,
              annotation_text=f"Pipe goal: {al_g_0:,.0f}, Δ: {al_d_0:,.0f} ",
              annotation_position="top right",
              line_color="#C28800")
#Fabris 3° mese
fig.add_hline(y=cf_g_0, line_dash="dot", row=3, col=3,
              annotation_text=f"Pipe goal: {cf_g_0:,.0f}, Δ: {cf_d_0:,.0f} ",
              annotation_position="top right",
              line_color="#C28800")
#Mozzolani 3° mese
fig.add_hline(y=cm_g_0, line_dash="dot", row=3, col=4,
              annotation_text=f"Pipe goal: {cm_g_0:,.0f}, Δ: {cm_d_0:,.0f} ",
              annotation_position="top right",
              line_color="#C28800")


#fig.update_yaxes(showticklabels=True, visible=True)

#fig.add_annotation(row=3, col=1,text="Ciao", x= "Commit Forecast", y=500000, hovertext="ciao",bgcolor="red", borderpad=3)

# hide subplot y-axis titles and x-axis titles
for axis in fig.layout:
    if type(fig.layout[axis]) == go.layout.YAxis:
        fig.layout[axis].title.text = ''
    if type(fig.layout[axis]) == go.layout.XAxis:
        fig.layout[axis].title.text = ''

fig.show()

pyo.plot(fig, filename=r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Salesforce\pipe.html')

import win32com.client as client

dashboard_inflow_indiretta = r'C:\Users\Raffaele.Sportiello\OneDrive - Wolters Kluwer\Documents\Salesforce\pipe.html'

html_body = """
    <div>
          <p>Ciao a tutti,<br><br>
            in allegato il file aggiornato.<br><br>
            Un saluto,<br>Raffaele<br><br></p>
    </div>
"""


outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0)
message.To = "Andrea.Ferrara@wolterskluwer.com;Camilla.Fabris@wolterskluwer.com;Aureliano.Leone@wolterskluwer.com;Cristiano.Mozzolani@wolterskluwer.com"
message.CC = "marco.bitossi@wolterskluwer.com;claudio.ferrante@wolterskluwer.com"
message.Subject = 'Inflow indiretta'
message.HTMLBody = html_body
message.Attachments.Add(Source=dashboard_inflow_indiretta)

message.Display()
