import json
import clr # From pythonnet, remove any existing clr -- pip uninstall clr
import os


folder = r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"

clr.AddReference(folder + 
    r"\Microsoft.AnalysisServices\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.dll")

clr.AddReference(folder +
    r"\Microsoft.AnalysisServices.Tabular\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll")

clr.AddReference(folder +
    r"\Microsoft.AnalysisServices.Tabular.Json\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.Json.dll")


import Microsoft.AnalysisServices as AS
import Microsoft.AnalysisServices.Tabular as Tabular


workspace_xmla = "source workspace XMLA"
username = 'pbi service user id'
password = 'pbi sevice user pass'
conn_string = f"DataSource={workspace_xmla};User ID={username};Password={password};"

server = Tabular.Server()
server.Connect(conn_string)

folder_path = r"C:\Users\antsharma\Downloads\Power BI\\"
  

def export_model_json(server: Tabular.Server):
    
    for db in server.Databases:

        script = Tabular.JsonScripter.ScriptCreate(db)
        json_file = json.loads(script)['create']['database']
        raw_json = json.dumps(json_file, indent=4)
        
        with open(folder_path + db.Name + '.bim', 'w') as model_bim:
            model_bim.write(raw_json)
            
            
export_model_json(server)
server.Disconnect()

# Second step is to read the file and generate the database:

workspace_xmla = "new Power BI workspace"
username = 'pbi service user id'
password = 'pbi sevice user pass'
conn_string = f"DataSource={workspace_xmla};User ID={username};Password={password};"

server = Tabular.Server()
server.Connect(conn_string)


def publish_model_bim(bim_file_path, server: Tabular.Server):
    
    for filename in os.listdir(bim_file_path):
        f = os.path.join(directory, filename)
        
        if os.path.isfile(f):
            file_name = os.path.splitext(os.path.basename(f))[0]
            new_dataset_name = server.Databases.GetNewName(file_name)
            
            with open(f) as bim:
                json_file = json.load(bim)
                #json_file.update({'compatibilityLevel':1571})
                json_file['id'] = new_dataset_name
                json_file['name'] = new_dataset_name
                json_file['model']['defaultPowerBIDataSourceVersion'] = "powerBI_V3"
            
            raw_json = json.dumps(json_file, indent = 4)

            db = AS.JsonSerializer.DeserializeDatabase(
                raw_json, 
                DeserializeOptions = 'default', 
                CompatibilityMode = 'PowerBI'
            )

            script = Tabular.JsonScripter.ScriptCreateOrReplace(db)
            server.Execute(script)

            
publish_model_bim(folder_path, server)
server.Disconnect()
