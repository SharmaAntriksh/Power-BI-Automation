import msal
import requests
import pandas as pd
import itertools


def get_access_token():
    
    app_id = '*****************************************'
    pbi_tenant_id = ''*****************************************''
    app_secret = ''*****************************************''
    
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
    
    # workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups?$filter=(isOnDedicatedCapacity eq true)'
    workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups'
    response_request = requests.get(workspace_endpoint, headers=headers)
    result = response_request.json()
    
    workspace_id = [workspace['id'] for workspace in result['value']]
    workspace_name = [workspace['name'] for workspace in result['value']]
    
    return zip(workspace_id, workspace_name)
    
    
def get_workspace_users():
    
    workspace_ids, workspace_names = zip(*list(get_workspaces()))
    user_details = []

    for workspace_id in workspace_ids:
        users_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/users"
        response = requests.get(users_endpoint, headers = headers)
        response_text = response.json()['value']
        for entry in response_text:
            entry['WorkspaceID'] = workspace_id
        
        user_details.append(response_text)
        
    return list(itertools.chain(*user_details))
    
    
workspace_df = pd.DataFrame.from_records(get_workspaces(), columns = ['WorkspaceID', 'WorkspaceName'])
workspace_user_df = pd.DataFrame(get_workspace_users())

final_df = pd.merge(workspace_df, workspace_user_df, how = 'inner', on = 'WorkspaceID' )
final_df.rename(
    columns = {
        'WorkspaceID': 'Workspace ID',
        'WorkspaceName': 'Workspace Name',
        'groupUserAccessRight': 'User Access',
        'displayName': 'User Name',
        'identifier': 'User ID',
        'principalType': 'User Type',
        'emailAddress': 'Email ID'
    },
    inplace = True
)

final_df['Email ID'] = final_df['Email ID'].fillna('')
final_df.to_excel(r"C:\Users\antsharma\OneDrive\Desktop\User Details.xlsx", index = False)
