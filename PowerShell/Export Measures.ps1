[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.Tabular")

$Server = New-Object Microsoft.AnalysisServices.Tabular.Server

$ServerName = "localhost:62465"
$DatabaseName = "a121f7d4-6e5f-4ce9-b806-f01ce6335c04" # Retrieved using DAX Studio
$UserID = "domain\username"
$UserPass = "**************"

$ConnectionString = "DataSource=$($ServerName);Initial Catalog=$($DatabaseName);User ID=$($UserID);Password=$($UserPass);"
$Server.Connect($ConnectionString)

# $Model = $Server.Databases.FindByName("Contoso Import").Model  # For SSAS
$Model = $Server.Databases[0].Model # For Power BI Desktop

$Tables = New-Object System.Collections.Generic.List[Microsoft.AnalysisServices.Tabular.Table]
$MeasuresList = New-Object System.Collections.Generic.List[Microsoft.AnalysisServices.Tabular.Measure] 

$Tables = $Model.Tables | ForEach-Object {
    $Table = $_
    $Measures = $Table.Measures | ForEach-Object { 
        $Measure = $_
        $MeasuresList += $Measure
    }
}

$MeasuresList |
    ` Select-Object -Property Name, Description, DataType, Expression, FormatString, IsHidden, State, ModifiedTime, StructureModifiedTime, IsSimpleMeasure |
    ` Export-Csv $ENV:USERPROFILE\Downloads\Measures.csv
