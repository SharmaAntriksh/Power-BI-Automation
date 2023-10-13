import msal
import requests

app_id = '********************************'
tenant_id = '********************************'
app_secret = '********************************'

user_id = "abc@123.onmicrosoft.com"
user_pass = '********************************'


def get_access_token_user():
    
    authority_url = f'https://login.microsoftonline.com/{tenant_id}'
    scopes = [r'https://analysis.windows.net/powerbi/api/.default']

    client = msal.PublicClientApplication(app_id, authority=authority_url)
    response = client.acquire_token_by_username_password(user_id, user_pass, scopes)
    access_id = response.get('access_token')

    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {access_id}'
    }

    return headers

user_headers = get_access_token_user()


def get_workspaces():
    
    workspace_endpoint = 'https://api.powerbi.com/v1.0/myorg/groups?$filter=(isOnDedicatedCapacity eq true)'
    response_request = requests.get(workspace_endpoint, headers=user_headers)
    result = response_request.json()
    
    workspace_id = [workspace['id'] for workspace in result['value']]
    workspace_name = [workspace['name'] for workspace in result['value']]
    
    return zip(workspace_id, workspace_name)


workspaces = get_workspaces()


def get_datasets(workspace_id = None):
    
    if workspace_id is None:
        all_workspaces = workspaces
    else:
        all_workspaces = filter(lambda ws: ws[0] == workspace_id, workspaces)

    dataset_ids = []
    dataset_names = []
    workspace_ids = []

    for workspace_id, workspace_name in all_workspaces:
        dataset_endpoint = fr"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets"
        response_request = requests.get(dataset_endpoint, headers=user_headers)
        result = response_request.json()

        for item in result['value']:
            dataset_ids.append(item['id'])
            dataset_names.append(item['name'])
            workspace_ids.append(workspace_id)

    return zip(dataset_ids, dataset_names, workspace_ids)


def take_over_dataset(workspace_id):
    
    datasets = get_datasets(workspace_id)
    
    for dataset_id, dataset_name, workspace_id in datasets:
        take_over = fr'https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/Default.TakeOver'
        response_request = requests.post(take_over, headers=user_headers)
        if response_request.status_code == 200:
            # result = response_request.json() 
            # Currently it looks like there is no body of the response
            print(f'Take over of dataset: {dataset_name} complete')


take_over_dataset('********************************')


def bind_dataset_gateway(workspace_id, gateway_id):

    gateway_details = {
        'gatewayObjectId': gateway_id
    }

    datasets = get_datasets(workspace_id)

    for dataset_id, dataset_name, workspace_id in datasets:
        bind = fr'https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{dataset_id}/Default.BindToGateway'
        response_request = requests.post(bind, headers=user_headers, json=gateway_details)
        if response_request.status_code == 200:
            print(f'Binding complete for dataset {dataset_name}.')

bind_dataset_gateway('********************************', '********************************')
