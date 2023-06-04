# Install-Module -Name MicrosoftPowerBIMgmt

$AppId = "924418cd-******************************"
$TenantId = "443c0ee8-***************************"
$ClientSecret = "cUh8Q~**************************" 

$SecurePassword = ConvertTo-SecureString $ClientSecret -Force -AsPlainText
$Credential = New-Object Management.Automation.PSCredential($AppId, $SecurePassword)

Connect-PowerBIServiceAccount -ServicePrincipal -TenantId $TenantId -Credential $Credential
$Headers = Get-PowerBIAccessToken 

$Workspaces = Get-PowerBIWorkspace -Scope Individual | 
    Where-Object { 
        $_.IsOnDedicatedCapacity -eq 'true' -and $_.Name -notin ('Admin monitoring', 'Fabric Demo') 
    }

$ResultArray = @()
$ApiString = "https://api.powerbi.com/v1.0/myorg/"

$Workspaces | ForEach-Object {

    $Workspace = $_
    $Datasets = Invoke-RestMethod `
                    -Method Get `
                    -Headers $Headers `
                    -Uri "$($ApiString)/groups/$($Workspace.Id)/datasets" 

    $Datasets.Value | ForEach-Object {

        $Dataset = $_
        $Refreshes = Invoke-RestMethod `
                        -Method Get `
                        -Headers $Headers `
                        -Uri  "$($ApiString)/groups/$($Workspace.Id)/datasets/$($Dataset.Id)/refreshes"

        $Refreshes.Value | ForEach-Object {

            $Refresh = $_
            $Row = New-Object PSObject -Property @{
                WorkspaceID      = $Workspace.Id;
                WorkspaceName    = $Workspace.Name;
                DatasetID        = $Dataset.Id;
                DatasetName      = $Dataset.Name;
                RefreshID        = $Refresh.RequestId;
                RefreshType      = $Refresh.refreshType;
                RefreshStartTime = $Refresh.startTime;
                RefreshEndTime   = $Refresh.endTime
            }

            $ResultArray += $Row
        }
    }
    
}

$ResultArray | 
    Select-Object -Property `
        WorkspaceID, WorkspaceName, DatasetID, DatasetName, 
        RefreshID, RefreshType, RefreshStartTime, RefreshEndTime |

    Export-Csv -NoTypeInformation `
        -Path "$($env:USERPROFILE)\Downloads\refresh.csv"
