
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.Tabular")

$FileName = "C:\Users\antsharma\OneDrive\Desktop\PBI Service Credentials.txt"

$Credentials = Get-Credential
$Credentials | Export-Clixml -Path $FileName
$FileCred = Import-Clixml -Path $FileName

$UserName = $FileCred.UserName
$UserPass = [System.Net.NetworkCredential]::new("", $FileCred.Password).Password

Import-CSV "C:\Users\antsharma\OneDrive\Desktop\Power BI Refresh Schedule.csv" | 
   ` Select-Object WorkspaceName, DatasetID, TableName, DatasetName |
   ` ForEach-Object {
        $WorkspaceName = $_.WorkspaceName
        $TableName = $_.TableName
        $DatasetName = $_.DatasetName

        $ServerXMLA = "powerbi://api.powerbi.com/v1.0/myorg/$($WorkspaceName)"
        $Server = New-Object Microsoft.AnalysisServices.Tabular.Server

        $ConnectionString = "DataSource=$($ServerXMLA);User ID=$($UserName);Password=$($UserPass)"
        $Server.Connect($ConnectionString)

        [Microsoft.AnalysisServices.Tabular.Database]$Database = $Server.Databases.FindByName("$($DatasetName)")
        [Microsoft.AnalysisServices.Tabular.Model]$Model = $Database.Model

        [Microsoft.AnalysisServices.Tabular.Table]$TargetTable = $Model.Tables["$($TableName)"]
        $TargetTable.RequestRefresh("Full")

        $Model.SaveChanges()
        $Server.Disconnect()
     }
      
