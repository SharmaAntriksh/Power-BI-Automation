import json
import clr # From pythonnet

folder = r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"

clr.AddReference(folder + 
    r"\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.DLL")

clr.AddReference(folder +
    r"\Microsoft.AnalysisServices\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.DLL")

import Microsoft.AnalysisServices as AS
import Microsoft.AnalysisServices.Tabular as Tabular

workspace = "Domain\SSAS InstanceName"
username = 'domain\WindowsAccountName'
password = ''
conn_string = f"DataSource={workspace};User ID={username};Password={password};"

server = Tabular.Server()
server.Connect(conn_string)
new_dataset_name = server.Databases.GetNewName('SSAS Tabular Model w Python')

bim_file = r"C:\Users\antsharma\Downloads\Model.bim"

with open(bim_file) as bim:
    json_file = json.load(bim)
    json_file['id'] = new_dataset_name
    json_file['name'] = new_dataset_name

raw_json = json.dumps(json_file, indent=4)

db = AS.JsonSerializer.DeserializeDatabase(
    raw_json, 
    DeserializeOptions = 'default', 
    CompatibilityMode = 'Analysis Services'
)
# Compatibility Modes:
    # Analysis Services = 1 (Basic AnalysisServices mode - used on SSAS and AAS)
    # Excel PowerPivot = 4
    # PowerBI = 2
    # Unknown = 0

script = Tabular.JsonScripter.ScriptCreateOrReplace(db)
server.Execute(script)
server.Disconnect()
