import msal
import requests
import pandas as pd
import itertools

# pip install msrest, mstrestazure
# pip install azure.graphrbac
# pip install azure.common

from azure.common.credentials import ServicePrincipalCredentials
from azure.graphrbac import GraphRbacManagementClient

app_id = '*********************************************'
pbi_tenant_id = '*********************************************'
app_secret = '*********************************************'


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


def get_workspaces():
    
    # workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups?$filter=(isOnDedicatedCapacity eq true)'
    
    workspaces = 'https://api.powerbi.com/v1.0/myorg/groups'
    response_request = requests.get(workspaces, headers=headers)
    result = response_request.json()
    
    workspace_id = [workspace['id'] for workspace in result['value']]
    workspace_name = [workspace['name'] for workspace in result['value']]
    
    return zip(workspace_id, workspace_name)


def get_group_members(group_id):
    
    # Code taken from https://stackoverflow.com/questions/51859504/how-to-access-the-azure-ad-groups-and-user-details-using-python
    
    credentials = ServicePrincipalCredentials(
        client_id = app_id,
        secret = app_secret,
        resource = "https://graph.windows.net",
        tenant = '16ky80.onmicrosoft.com' # Change 16ky80 to your tenant
    )

    graphrbac_client = GraphRbacManagementClient(
        credentials,
        pbi_tenant_id
    )
    
    users = graphrbac_client.groups.get_group_members(group_id)
    
    result = ''
    
    for u in users:
        if u.object_type == 'User':
            result += u.mail  + '\n'
        elif u.object_type == 'ServicePrincipal':
             result +=  'Service Principal - ' + u.display_name + '\n'

    return result[:-1]


def get_workspace_users():
    
    workspace_ids, workspace_names = zip(*list(get_workspaces()))
    user_details = []

    for workspace_id in workspace_ids:
        users = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/users"
        response = requests.get(users, headers = headers)
        response_text = response.json()['value']
        
        for entry in response_text:
            entry['WorkspaceID'] = workspace_id
            if entry['principalType'] == 'Group':
                entry['Group Membership'] = get_group_members(entry['identifier'])
        
        user_details.append(response_text)
        
    return list(itertools.chain(*user_details))


workspace_df = pd.DataFrame.from_records(get_workspaces(), columns = ['WorkspaceID', 'WorkspaceName'])
workspace_user_df = pd.DataFrame(get_workspace_users())

final_df = pd.merge(workspace_df, workspace_user_df, how = 'left', on = 'WorkspaceID' )
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
final_df.to_excel(r"C:\Users\antsharma\OneDrive\Desktop\User Details.xlsx", index = False)
