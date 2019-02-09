 param (
    [string]$SmartsheetId 
 )
# Documentation Notes
# Explain how to get API key from Smartsheet

function Write-SQL {
    $updateSQLCheck = ""
    $updateSQL = ""
    $insertColumns = ""
    $insertValues = ""
    foreach($column in $columnList) {
        if (!($column.title -eq "SmartsheetId")) {
            #UpdateSQLCheck
            if ($updateSQLCheck -eq "") {
                $updateSQLCheck = "TARGET.[$($column.dbtitle)] <> SOURCE.[$($column.dbtitle)]"
            }
            else {
                $updateSQLCheck += " OR TARGET.[$($column.dbtitle)] <> SOURCE.[$($column.dbtitle)]"
            }

            #UpdateSQL
            if ($updateSQL -eq "") { $updateSQL = "TARGET.[$($column.dbtitle)] = SOURCE.[$($column.dbtitle)]" }
            else { $updateSQL += ", TARGET.[$($column.dbtitle)] = SOURCE.[$($column.dbtitle)]" }
        
            #InsertColumns
            if ($insertColumns -eq "") { $insertColumns = "[$($column.dbtitle)]" }
            else { $insertColumns += ", [$($column.dbtitle)]" }

            #InsertValues
            if ($insertValues -eq "") { $insertValues = "SOURCE.[$($column.dbtitle)]" }
            else { $insertValues += ", SOURCE.[$($column.dbtitle)]" }
        }
    }

    $targetTable = $config.dbTableName
    $targetID = "SmartsheetId"
    $sourceTable = $config.dbTempTableName
    $sourceID = "SmartsheetId"

    ### Generate SQL Merge ###   
    $sql = @"
    MERGE $targetTable AS TARGET
    USING $sourceTable AS SOURCE 
    ON (TARGET.[SmartsheetId] = SOURCE.[SmartsheetId]) 
    --Update existing records
    WHEN MATCHED AND $updateSQLCheck 
    THEN UPDATE SET $updateSQL
    --Insert new reocrds
    WHEN NOT MATCHED BY TARGET THEN 
    INSERT ([SmartsheetId],$insertColumns) 
    VALUES (SOURCE.[SmartsheetId],$insertValues)
    --Delete records not in source
    WHEN NOT MATCHED BY SOURCE THEN 
    DELETE;
"@
    
    $sql | Out-File -FilePath "$PSScriptRoot\SmartsheetSync-$($config.smartsheetID).sql"
}

# Show config menu
Clear-Host
Write-Host "================== Generate Smartsheet sync configuration ==================="
Write-Host "This script helps you prepare for a Smartsheet -> SQL Server process"
Write-Host "Please answer the following questions to generate the necessary configuration"
Write-Host "============================================================================="

# Load existing config (or use blank one)
$configFile = "$PSScriptRoot\config-$SmartsheetId.json"
if (!(Test-Path -Path $configFile)) {
    $configFile = "$PSScriptRoot\config-default.json"
}    
$config = Get-Content -Raw -Path $configFile | ConvertFrom-Json

#Get Smartsheet API Key
$defaultValue = $config.bearer
$prompt = Read-Host "Please enter the API key for Smartsheet" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.bearer = $prompt

#Get Smartsheet ID
$defaultValue = $config.smartsheetID
$prompt = Read-Host "Please enter the ID of the Smartsheet you want to sync (eg: 8534977916753796)" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.smartsheetID = $prompt

#Connect via REST API call
Write-Host "Connecting to Smartsheet..." -ForegroundColor Green
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization","Bearer " + $config.bearer)
$headers.Add("Content-Type","application/json")
$jsondata = Invoke-RestMethod https://api.smartsheet.com/2.0/sheets/$($config.smartsheetID)/ -Method Get -Headers $headers
$config.smartsheetName = $jsondata.name
Write-Host -NoNewLine "Connected to " -ForegroundColor Green
Write-Host -NoNewLine $ssname -ForegroundColor Cyan
Write-Host " Smartsheet" -ForegroundColor Green

#Get Column Details
Write-Host "Retrieving Smartsheet columns..." -ForegroundColor Green
$columnList = @()
$columnIndex = 1
foreach($column in $jsondata.columns)
{
    $columnProp = [ordered]@{}
    $columnProp.add('Index', $columnIndex)
    $columnProp.add('Title', $column.title)
    $columnProp.add('DBTitle', $column.title.Replace(" ", "").Replace("-","_").Replace("(","_").Replace(")","_"))
    $columnProp.add('Type', $column.type)
    $columnProp.add('Sync', 'true')
    $columnObj = New-Object -TypeName psobject -Property $columnProp
    $columnList += $columnObj
    $columnIndex++
}
#$columnList | Format-Table -AutoSize @{L='ID';E={$_.Index}}, Title
#[int]$readinput = Read-Host "Which column should be the primary key? Enter the ID"
#$config.smartsheetUniqueID = ($columnList |?{ $_.Index -eq $readinput}).Title 
#Write-Host "You selected " -NoNewLine -ForegroundColor Green
#Write-Host $config.smartsheetUniqueID -ForegroundColor Cyan 
$config.columns = $columnList

#Set Database table names
$ssname = $config.smartsheetName.Replace(' ', '_')
$dbTableName = "SS_" + $ssname
$defaultValue = $dbtablename
$prompt = Read-Host "What would you like to name the database table? Press <Enter> to accept the default [$($defaultValue)]"
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.dbTableName = $prompt

#Set Database temp table name
$defaultValue = $config.dbTableName + '_temp'
$prompt = Read-Host "What would you like to name the database temp table? Press <Enter> to accept the default [$($defaultValue)]"
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.dbTempTableName = $prompt

#Set Database Server name
$defaultValue = $config.connectionstring.server
$prompt = Read-Host "Please enter the SQL Server hostname" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.connectionstring.server = $prompt

#Set Database Table name
$defaultValue = $config.connectionstring.database
$prompt = Read-Host "Please enter the database name" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
$config.connectionstring.database = $prompt

#Set Integrated Authentication
$defaultValue = $config.connectionstring.useSSP
$prompt = Read-Host "Would you like to use integrated authentication (Windows login) to access the database [Y/N]" 
$prompt = ($defaultValue,$prompt)[[bool]$prompt]
if ($prompt -eq "Y") { 
    $config.connectionstring.useSSP = "true" 
    $config.connectionstring.user = ""
    $config.connectionstring.password = ""
}
else { 
    $config.connectionstring.useSSP = "false" 

    #Set database username
    $defaultValue = $config.connectionstring.user
    $prompt = Read-Host "Please enter the database username" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
    $prompt = ($defaultValue,$prompt)[[bool]$prompt]
    $config.connectionstring.user = $prompt

    #Set database user password
    $defaultValue = $config.connectionstring.password
    $prompt = Read-Host "Please enter the database password" $(If ($defaultValue -eq "") {""} Else {". Press <Enter> to accept current value [$($defaultValue)]"}) 
    $prompt = ($defaultValue,$prompt)[[bool]$prompt]
    $config.connectionstring.password = $prompt
}

#Create Scheduled Task
$scheduledTask = Read-Host "Would you like to create a scheduled task to perform the sync? [Y/N]"
if ($scheduledTask -eq "Y") {
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' `
      -Argument "-File '$PSScriptRoot\SmartsheetToSQL.ps1' -SmartsheetId $($config.smartsheetID)"
    $trigger =  New-ScheduledTaskTrigger -Daily -At 5am
    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "SmartsheetSync" -Description "Daily sync of Smartsheet" | out-null
    Write-Host -NoNewline "Created scheduled task named " -ForegroundColor Green
    Write-Host -NoNewLine "SmartsheetSync " -ForegroundColor Cyan
    Write-Host "to run daily at 5am. Please review documentation to configure the scheduled task as needed" -ForegroundColor Green
}


#Generate config file
Write-Host "Generating Smartsheet sync configuration file..." -ForegroundColor Green
$config | ConvertTo-Json | Out-File -FilePath "$PSScriptRoot\config-$($config.smartsheetID).json"
Write-Host "Completed Smartsheet sync configuration file: " -NoNewline -ForegroundColor Green
Write-Host "config-$($config.smartsheetID).json" -ForegroundColor Cyan

#Generate SQL script file
Write-Host "Generating SQL script file..." -ForegroundColor Green
Write-SQL
Write-Host "Completed SQL script file: " -NoNewline -ForegroundColor Green
Write-Host "SmartsheetSync-$($config.smartsheetID).sql" -ForegroundColor Cyan



#Write out instructions
Write-Host "-------------------------------------------------------" -ForegroundColor White
Write-Host "To perform the Smartsheet sync, run the following script:" -ForegroundColor Green
Write-Host ".\SmartsheetToSQL.ps1 -SmartsheetID $($config.smartsheetID)" -ForegroundColor Yellow 

