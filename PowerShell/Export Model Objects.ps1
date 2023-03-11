#Work in Progress

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.Tabular")

$Server = New-Object Microsoft.AnalysisServices.Tabular.Server

$ServerName = "powerbi://api.powerbi.com/v1.0/myorg/Incremental%20Refresh%20Demo"
$DatabaseName = "a121f7d4-6e5f-4ce9-b806-f01ce6335c04" # Retrieved using DAX Studio
$UserID = ""
$UserPass = ""

$ConnectionString = "DataSource=$($ServerName);User ID=$($UserID);Password=$($UserPass);"
$Server.Connect($ConnectionString)

# $Model = $Server.Databases.FindByName("Contoso Import").Model  # For SSAS
# $Model = $Server.Databases[0].Model # For Power BI Desktop
$Model = $Server.Databases.FindByName("Contoso 100K").Model

$TablesList = New-Object System.Collections.ArrayList
$MeasuresList = New-Object 'System.Collections.Generic.List[Tuple[Microsoft.AnalysisServices.Tabular.Measure, string]]'
$ColumnsList = New-Object System.Collections.ArrayList

$Model.Tables | ForEach-Object {
    $Table = $_
    $TablesList += $Table
    $Table.Measures | ForEach-Object { $MeasuresList.Add([Tuple]::Create($_, $Table.Name)) }
    $Table.Columns | ForEach-Object { [void]$ColumnsList.Add([Tuple]::Create($_, $Table.Name)) }
}

$ExcelApp = New-Object -comobject Excel.Application
$ExcelApp.Visible = $true

# Table Worksheet

$Workbook = $ExcelApp.Workbooks.Add()
$Worksheet = $Workbook.Worksheets[1]
$Worksheet.Name = "Tables"

$ColumnNames = @(
    'Measure Name', 
    "Expression", 
    "Format String", 
    "Hidden", 
    "Data Type"
)

for($col = 1; $col -le $ColumnNames.length;  $col++){
    $WorkSheet.Cells.Item(1,$col) = $ColumnNames[$col - 1]
}

$Model.Tables | ForEach-Object {
    $Table = $_
    #$Table.Measures | ForEach-Object { $MeasuresList.Add([Tuple]::Create($_, $Table.Name)) }
    $Table.Columns | ForEach-Object { $ColumnsList.Add([Tuple]::Create($_, $Table.Name)) }

    $TablesList
