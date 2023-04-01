import msal
import requests
import pandas as pd
from pathlib import Path


def get_access_token():
    
    app_id = '*********************************************'
    pbi_tenant_id = '*********************************************'
    app_secret = '*********************************************'
    
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
    
    workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups?$filter=(isOnDedicatedCapacity eq true)'
    response_request = requests.get(workspace_endpoint, headers=headers)
    result = response_request.json()
    
    workspace_id = [workspace['id'] for workspace in result['value']]
    workspace_name = [workspace['name'] for workspace in result['value']]
    
    return zip(workspace_id, workspace_name)


def get_datasets(workspace_id = None):
    
    if workspace_id is None:
        all_workspaces = get_workspaces()
    else:
        all_workspaces = [workspace_id]

    dataset_ids = []
    dataset_names = []
    workspace_ids = []

    for workspace_id, workspace_name in all_workspaces:

        dataset_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
        response_request = requests.get(dataset_endpoint, headers=headers)
        result = response_request.json()

        for item in result['value']:
            dataset_ids.append(item['id'])
            dataset_names.append(item['name'])
            workspace_ids.append(workspace_id)

    return zip(dataset_ids, dataset_names, workspace_ids)


def get_refresh_history():
    
    dataset_ids, dataset_names, workspace_ids = zip(*list(get_datasets()))
    dataset_workspace = zip(dataset_ids, workspace_ids)
    result = []
    
    for dataset_id, workspace_id in dataset_workspace:
        refresh_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
        response_request = requests.get(refresh_endpoint, headers=headers)
        
        if response_request.status_code == 200:
            response_text = response_request.json()['value']
            for entry in response_text:
                entry['DatasetID'] = dataset_id
                entry['WorkspaceID'] = workspace_id
            
            result.append(response_text)
    
    flatlist = []

    for sublist in result:
        for element in sublist:
            flatlist.append(element)
            
    return flatlist


workspace_df = pd.DataFrame.from_records(get_workspaces(), columns = ['WorkspaceID', 'WorkspaceName'])
dataset_df = pd.DataFrame.from_records(get_datasets(), columns = ['DatasetID', 'DatasetName', 'WorkspaceID'])

refresh_df = pd.DataFrame(get_refresh_history())
columns_to_keep = ['refreshType', 'startTime', 'endTime', 'status', 'DatasetID', 'WorkspaceID', 'serviceExceptionJson']
capitalize_names = [name[0].capitalize() + name[1:] for name in columns_to_keep]
refresh_df = refresh_df[columns_to_keep]
refresh_df.columns = capitalize_names
refresh_df = refresh_df.astype({'StartTime': 'datetime64[ns]', 'EndTime': 'datetime64[ns]'})
refresh_df['StartTime'] = refresh_df['StartTime'].dt.date
refresh_df['EndTime'] = refresh_df['EndTime'].dt.date

final_df = pd.merge(workspace_df, dataset_df, how='inner', left_on = 'WorkspaceID', right_on = 'WorkspaceID').merge(refresh_df, how = 'inner', on = 'WorkspaceID')
columns_to_keep = ['WorkspaceID', 'WorkspaceName', 'DatasetID_x', 'DatasetName', 'RefreshType', 'StartTime', 'EndTime', 'ServiceExceptionJson']
final_df = final_df[columns_to_keep]
final_df.rename(columns = {'DatasetID_x':'DatasetID'}, inplace = True)

home = str(Path.home())
final_df.to_excel(home + r"\Desktop\Refresh History.xlsx", index= False)
