 param (
    [string]$SmartsheetId 
 )

#-----------------------------------------------------------
#-------------------- SQL Functions ------------------------
#-----------------------------------------------------------
function Test-SQLConnection
{    
    [OutputType([bool])]
    Param
    (
        [Parameter(Mandatory=$true,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0)]
        $ConnectionString
    )
    try
    {
        $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $ConnectionString;
        $sqlConnection.Open();
        $sqlConnection.Close();

        return $true;
    }
    catch
    {
        return $false;
    }
}

function Invoke-SqlCommand {
    Param (
        [Parameter(Mandatory = $true)][string]$Query = $(throw "Please specify a query."),
        [Parameter(Mandatory = $true)][string]$Connectionstring = $(throw "Please specify a connection string")
    )

    #connect to database
    $connection = New-Object System.Data.SqlClient.SqlConnection($Connectionstring)
    $connection.Open()
    
    #build query object
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $command.CommandTimeout = $CommandTimeout
    
    #run query
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    [void]$adapter.Fill($dataset) #| out-null
    
    #return the first collection of results or an empty array
    If ($null -ne $dataset.Tables[0]) { $table = $dataset.Tables[0] }
    ElseIf ($table.Rows.Count -eq 0) { $table = New-Object System.Collections.ArrayList }
    
    $connection.Close()
    return , $table
}


#Build the tables if they don't exist and truncate temp table
function ResetTables {
    $tempTable = Invoke-SqlCommand -Connectionstring $cfgConnString -Query "SELECT CASE WHEN OBJECT_ID('dbo.$($config.dbTempTableName)', 'U') IS NULL THEN 'False' ELSE 'True' END AS TableExists"
    $table = Invoke-SqlCommand -Connectionstring $cfgConnString -Query "SELECT CASE WHEN OBJECT_ID('dbo.$($config.dbTableName)', 'U') IS NULL THEN 'False' ELSE 'True' END AS TableExists"

    #ADD IF LOGIC HERE TO SEE IF TABLE EXISTS OR NOT
    if ("False" -eq $table.Rows[0].TableExists) {
        #Read data from config file for columns
        $columnList=""
        foreach($column in $config.columns)
        {

            $columnTitle = $column.dbtitle
            $columnType = $column.type
            $columnId = $column.id

            #Generate table schema
            if ($column.sync) {
                if($columnType -eq "PICKLIST")
                {
                    $columnList+= $columnTitle + " nvarchar(1000),"
                }
                elseif($columnType -eq "DATE")
                {
                    #removes any spaces from the column names. Sets 1000 character limit in SQL, but could be increased
                    $columnList+= $columnTitle + " date,"
                }
                elseif($columnType -eq "DATETIME")
                {
                    #removes any spaces from the column names. Sets 1000 character limit in SQL, but could be increased
                    $columnList+= $columnTitle + " datetime,"
                }
                else
                {
                    $columnList+= $columnTitle + " nvarchar(1000),"
                }
            }
        }
        #removes the last caracter (the final comma)
        $columnList=$columnList -replace ".$"
    }
    
    #Create table
    if ("False" -eq $table.Rows[0].TableExists) {
        Write-Host "Creating table in database"
        $createTableSQL = "CREATE TABLE $($config.dbTableName) (SmartsheetId nvarchar(1000), $columnList)"
        Invoke-SqlCommand -Connectionstring $cfgConnString -Query $createTableSQL
    }

    #Create temp table
    if ("False" -eq $tempTable.Rows[0].TableExists) {
        Write-Host "Creating temp table in database"
        $createTempTableSQL = "CREATE TABLE $($config.dbTempTableName) (SmartsheetId nvarchar(1000), $columnList)"
        Invoke-SqlCommand -Connectionstring $cfgConnString -Query $createTempTableSQL
    }
}

#-----------------------------------------------------------
#------------------- Main Process --------------------------
#-----------------------------------------------------------
#Load configuration file for this sheet
Write-Host "Loading configuration file..."
$configFile = "$PSScriptRoot\config-$($config.smartsheetID).json"
if (!(Test-Path -Path $configFile)) {
    Write-Host "There is no configuration file for this Smartsheet. Please run the ConfigureSmartsheetSync.ps1 script before running this script"
    Return
}
$config = Get-Content -Raw -Path $configFile | ConvertFrom-Json

#Build connection string
$cfgServer = $config.connectionstring.server 
$cfgDatabase = $config.connectionstring.database
$cfgConnString = "Data Source=$cfgServer;Initial Catalog=$cfgDatabase;"
If ($config.connectionstring.useSSP -eq 'true') { $cfgConnString += "Integrated Security=SSPI;" } 
Else {
    $cfgDBUser = $config.connectionstring.user;
    $cfgDBPswd = $config.connectionstring.password; 
    $cfgConnString += "User ID=$username; Password=$password;" 
}

#Test database connection
Write-Host "Testing database connection..."
if (!(Test-SQLConnection $cfgConnString)) {
    Write-Host "Unabled to connect to database. Please check connection information and try again"
    Return
}

#Load MERGE SQL file
Write-Host "Loading SQL mapping file..."
$mergeFile = "$PSScriptRoot\SmartsheetSync-$($config.smartsheetID).sql"
if (!(Test-Path -Path $mergeFile)) {
    Write-Host "There is no SQL mapping file for this Smartsheet. Please run the ConfigureSmartsheetSync.ps1 script first before running this script"
    Return
}
$mergeSQL = Get-Content -Raw -Path $mergeFile 

#Reset Tables
Write-Host "Checking database tables..."
ResetTables

#Smartsheet REST API setup 
Write-Host "Connecting to Smartsheet..."
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization","Bearer $($config.bearer)")
$headers.Add("Content-Type","application/json")
$response = Invoke-RestMethod https://api.smartsheet.com/2.0/sheets/$($config.smartsheetId)/ -Method Get -Headers $headers

##Parse response
#$sheetname=$response.name
#$sheetname=$sheetname.Replace(' ', '_')

#Drop temp table
Invoke-SqlCommand -Connectionstring $cfgConnString -Query "IF OBJECT_ID('dbo.$($config.dbTempTableName)', 'U') IS NOT NULL TRUNCATE TABLE dbo.$($config.dbTempTableName)"

#Insert data to temp table
Write-Host "Copying Smartsheet data to temp table..."
#Loop over each row
ForEach($r in $response.rows)
{
    $id = $r.id
    
    #Build row data from cells
    $columnList=""
    ForEach($c in $r.cells)
    {
        if($c.value)
        {   
            #escapes any ' characters for SQL and truncates the string at 1000 characters
            $val = $c.value.ToString().Replace("'", "''")
            if($val.Length -gt 1000)
            {
                $val = $val.substring(0,1000)
            }
            $columnList+="'" + $val + "',"
        }
        else
        { 
            $columnList+="NULL,"
        }
    }
    $columnList=$columnList -replace ".$"
    
    $rowInsertSQL="INSERT INTO $($config.dbTempTableName) VALUES ('$id', $columnList);"
    Invoke-SqlCommand -Connectionstring $cfgConnString -Query $rowInsertSQL
}


#Execute MERGE SQL
Write-Host "Merging data into database..."
Invoke-SqlCommand -Connectionstring $cfgConnString -Query $mergeSQL

#Wrap-up
Write-Host "Smartsheet to SQL sync complete!"