import json
import clr # From pythonnet, remove any existing clr -- pip uninstall clr
import os
import msal
import requests


# Tabular Object Model DLLs that are required to interact with SSAS Dataset

folder = r"C:\Windows\Microsoft.NET\assembly\GAC_MSIL"

clr.AddReference(folder + 
    r"\Microsoft.AnalysisServices\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.dll")

clr.AddReference(folder +
    r"\Microsoft.AnalysisServices.Tabular\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll")

clr.AddReference(folder +
    r"\Microsoft.AnalysisServices.Tabular.Json\v4.0_19.61.1.4__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.Json.dll")

import Microsoft.AnalysisServices as AS
import Microsoft.AnalysisServices.Tabular as Tabular


app_id = '********************************************' # Azure application ID
pbi_tenant_id = '********************************************'
app_secret = '********************************************' # Secret code of that application


def get_access_token():
    
    authority_url = f'https://login.microsoftonline.com/{pbi_tenant_id}'
    scopes = [r'https://analysis.windows.net/powerbi/api/.default']

    client = msal.ConfidentialClientApplication(
        app_id, 
        authority=authority_url, 
        client_credential=app_secret
    )
    
    response = client.acquire_token_for_client(scopes)
    token = response.get('access_token')
    
    return token


token = get_access_token()
headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {token}'
}


# Get list of PPU workspaces as the TOM communication can happen only through XMLA
def get_workspaces():
    
    workspaces = 'https://api.powerbi.com/v1.0/myorg/groups?$filter=(isOnDedicatedCapacity eq true)'
    response_request = requests.get(workspaces, headers=headers)
    result = response_request.json()
    
    workspace_name = [workspace['name'] for workspace in result['value']]
    

# Write Downloads Model.bim file to a specific folder on system
def export_model_json(server: Tabular.Server):
    
    folder_path = r"C:\Users\antsharma\Downloads\Power BI\\"

    for db in server.Databases:

        script = Tabular.JsonScripter.ScriptCreate(db)
        json_file = json.loads(script)['create']['database']
        raw_json = json.dumps(json_file, indent=4)
        
        with open(folder_path + db.Name + '.bim', 'w') as model_bim:
            model_bim.write(raw_json)
            
            
# Iterate each workspace and downlod Model.bim file
def connect_workspace_and_export():
    
    workspaces = get_workspaces()
    
    for name in workspaces:
        workspace_xmla = f"powerbi://api.powerbi.com/v1.0/myorg/{name}"
        conn_string = f"DataSource={workspace_xmla};User ID=app:{app_id}@{pbi_tenant_id};Password={app_secret};"
        
        server = Tabular.Server()
        server.Connect(conn_string)
        
        export_model_json(server)
        server.Disconnect()
        

connect_workspace_and_export()


# Once all files are downloaded publish them to a workspace:

workspace_xmla = "powerbi://api.powerbi.com/v1.0/myorg/Demo%20PPU"
conn_string = f"DataSource={workspace_xmla};User ID=app:{app_id}@{pbi_tenant_id};Password={app_secret};"


def publish_model_bim(bim_file_path, server: Tabular.Server):
    
    for filename in os.listdir(bim_file_path):
        f = os.path.join(bim_file_path, filename)
        
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


publish_model_bim(r"C:\Users\antsharma\Downloads\Power BI\\", server)
server.Disconnect()
            server.Execute(script)
