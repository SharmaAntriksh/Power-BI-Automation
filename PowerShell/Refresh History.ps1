Login-PowerBIServiceAccount
$Headers = Get-PowerBIAccessToken

$Workspaces = Get-PowerBIWorkspace -Scope Individual | 
    ` Where-Object { 
            $_.IsOnDedicatedCapacity -eq 'true' -and $_.Name -notin ('Admin monitoring', 'Fabric Demo') 
    }

$ResultArray = @()

$Workspaces | ForEach-Object {
    $Workspace = $_

    $Datasets = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://api.powerbi.com/v1.0/myorg/groups/$($Workspace.Id)/datasets" 

    $Datasets.value | ForEach-Object {
        $Dataset = $_

        $Refreshes = Invoke-RestMethod -Method Get -Headers $Headers -Uri "https://api.powerbi.com/v1.0/myorg/groups/$($Workspace.Id)/datasets/$($Dataset.Id)/refreshes"

        $Refreshes.Value | ForEach-Object {
            $Refresh = $_

            $Row = New-Object PSObject -Property @{
                WorkspaceID = $Workspace.Id;
                WorkspaceName = $Workspace.Name;
                DatasetID = $Dataset.Id;
                DatasetName = $Dataset.Name;
                RefreshID = $Refresh.RequestId;
                RefreshType = $Refresh.refreshType;
                RefreshStartTime = $Refresh.startTime;
                RefreshEndTime = $Refresh.endTime
            }

            $ResultArray += $Row
        }
    }
    
}

$ResultArray | 
    ` Select-Object -Property WorkspaceID, WorkspaceName, DatasetID, DatasetName, RefreshID, RefreshType, RefreshStartTime, RefreshEndTime |
    ` Export-Csv -NoTypeInformation -Path "C:\Users\antsharma\Downloads\refresh.csv"
