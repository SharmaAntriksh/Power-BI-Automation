import clr # clr from pythonnet, must uninstall other clr and then install pythonnet
import pandas as pd

# if you have Power BI or Analysis Services installed then you will find this dll at C:\Windows\Microsoft.NET\assembly\GAC_MSIL

folder = r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
clr.AddReference(folder + r"\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.DLL")

import Microsoft.AnalysisServices.Tabular as Tabular

server = Tabular.Server()
connection_string = "localhost:56811" # Power BI

server.Connect(connection_string)

model  = server.Databases[0].Model #Power BI has only 1 database that's why 0 is used

# =========================================================================================
# For SSAS Tabular on premise or Power BI dataset use 

# SSAS :
  # workspace = "domain\servername"
  # username = 'domain\username'
  # password = '*****************'
  # conn_string = f"DataSource={workspace};User ID={username};Password={password};"
  # server.Connect(conn_string)

# Power BI dataset:
  # workspace = "powerbi://api.powerbi.com/v1.0/myorg/Your Workspace" # XMLA Endpoint
  # username = 'your login id form PBI Service'
  # password = 'your PBI Service password'
  # conn_string = f"DataSource={workspace};User ID={username};Password={password};"
  # server = Tabular.Server()
  # server.Connect(conn_string)

# =========================================================================================

all_measures = []

for table in model.Tables:
    for measure in table.Measures:
        all_measures.append(measure)

measures = {}
measure_names = []
measure_expressions = []

for m in all_measures:
    measure_names.append(m.Name)
    measure_expressions.append(m.Expression)

measures['Name'] = measure_names
measures['Expression'] = measure_expressions

df = pd.DataFrame(measures)
df.to_excel(r"C:\Users\antsharma\Downloads\Power BI Measures.xlsx", index = False)
