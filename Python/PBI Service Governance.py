import msal
import requests
import pandas as pd


def get_access_token():
    app_id = '924418cd-d505-4211-b223-343ade4aef26'
    pbi_tenant_id = '443c0ee8-0438-4a23-8872-ff8ceb3c4483'
    app_secret = 'cUh8Q~oBeH.abIoqIrkPYx4cvg0UH~MK.k_rGcl2'
    
    authority_url = f'https://login.microsoftonline.com/{pbi_tenant_id}'
    scopes = [r'https://analysis.windows.net/powerbi/api/.default']

    client = msal.ConfidentialClientApplication(app_id, authority=authority_url, client_credential=app_secret)
    
    response = client.acquire_token_for_client(scopes)
    token = response.get('access_token')
    
    return token
    

token = get_access_token()
headers = {
    'Content-Type': 'application/json',
    'Authorization': f'Bearer {token}'
}


def get_workspaces():
    workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups'
    response_request = requests.get(workspace_endpoint, headers=headers)
    result = response_request.json()
    
    workspace_id = [workspace['id'] for workspace in result['value']]
    workspace_name = [workspace['name'] for workspace in result['value']]
    
    return zip(workspace_id, workspace_name)
    
    
def get_datasets():
    all_workspaces = get_workspaces()
    workspace_datasets = {}

    for workspace_id, workspace_name in all_workspaces:
        
        dataset_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
        response_request = requests.get(dataset_endpoint, headers=headers)
        result = response_request.json()

        dataset_ids = [dataset['id'] for dataset in result['value']]
        dataset_names = [dataset['name'] for dataset in result['value']]
        
        workspace_datasets[workspace_name] = dataset_names

    return workspace_datasets
    

df = pd.DataFrame.from_dict(get_datasets(),orient='index')
df = df.transpose()
df.index += 1
df.fillna("",inplace=True)
df


def add_workspace_user():
    # Azure Group
    group_body = {
      "identifier": "29e3c3d2-b18e-46e9-92d2-dc03d2780a4e",
      "groupUserAccessRight": "Admin",
      "principalType": "Group"
    }
    
    # Individual User
    individual_user = {
      "emailAddress": "adelev@16ky80.onmicrosoft.com",
      "groupUserAccessRight": "Viewer"
    }

    workspace_ids, workspace_names = zip(*list(get_workspaces()))
    
    for workspace_id in workspace_ids:
        
        user_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/users"
        response_request = requests.post(user_endpoint, headers=headers, json = individual_user)
        
        if response_request.status_code == 200:
            print('user added succesfully')
        elif response_request.status_code == 404:
            result = response_request.json()
            print(result['error']['message'])
            
            
def remove_workspace_user():
    user = "adelev@16ky80.onmicrosoft.com"
    workspace_ids, workspace_names = zip(*list(get_workspaces()))
    
    for workspace_id in workspace_ids:
        
        user_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/users/{user}"
        response_request = requests.delete(user_endpoint, headers=headers)

        if response_request.status_code == 200:
            print('user removed succesfully')
        elif response_request.status_code == 404:
            result = response_request.json()
            print(result['error']['message'])
