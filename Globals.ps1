#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ConfigurationData
{
	$configData = (Get-Content .\config.json) | ConvertFrom-Json
	
	return $configData
}

function Log-ToFile
{
	param
	(
		[parameter(Mandatory = $true, Position = 1)]
		[string]$text,
		[parameter(Mandatory = $false,Position=2)]
		[string]$filename = $PWD.toString() + '\' + ((get-date -Format ddMMyyyy).toString()) + '.txt'
	)
	
	if(!(Test-Path ($filename)))
	{
		New-Item $filename -Type file
	}
	
	$row = (Get-Date -Format 'HH:mm:ss ddMMyyy') + ": " + $text
	
	$row | Out-File $filename -Append
}
function ConnectTo-SQLDatabase ()
{
	
	param (
		# Parameter help description
		[Parameter(Mandatory = $true, Position = 1)]
		$configuration
	)
	
	$dataSource = $configuration.databaseServer
	$database = $configuration.dbName
	
	#$authentication = "Provider=sqloledb";
	
	$authentication = "uid=" + $configuration.databaseUser + ";" +
	"pwd=" + $configuration.databaseUserPassword + ";"
	
	$connectionString = "Provider=sqloledb; " +
	"Data Source=$dataSource; " +
	"Initial Catalog=$database; " +
	"$authentication; "
	## Connect to the data source and open it
	
	Write-Host $connectionString
	
	$connection = New-Object System.Data.OleDb.OleDbConnection $connectionString
	$connection.Open()
	
	return $connection
}

function Import-AutoData ()
{
	param (
		[Parameter(Mandatory = $true, Position = 1)]
		$SQLConnection,
		[Parameter(Mandatory = $true, Position = 2)]
		$csvPath
	)

	$rows = Import-Csv $csvPath -Delimiter ","
	foreach	($row in $rows)
	{
        $sqlQuery = "INSERT INTO [dbo].[AUTO] ([NR_Zlecenia],[NR_Rejestracyjny],[TYP]) VALUES('{0}','{1}','{2}')" -f ($row.'Nr zlecenia').Trim(), ($row.'Numer rej').Trim(), ($row.TYP).Trim()
        $sqlCmd = New-Object System.Data.OleDb.OleDbCommand
        try
        {
            $sqlCmd.CommandText = $sqlQuery
            $sqlCmd.Connection = $SQLConnection
            if(($sqlCmd.ExecuteNonQuery()) -eq 1)
            {
                Log-ToFile "Object $($row.'Numer rej') added into database"
            }
        } 
        catch
        {
            if ($_.Exception.ErrorRecord.Exception.Message -like "*Cannot insert duplicate key in object*")
            {
                Log-ToFile "Object $($row.'Numer rej') already exists"
            }
            else 
            {
                Log-ToFile "Object $($row.'Numer rej') not added, unknown error"    
            }
        }
    }
}


#ConnectTo-SQLDatabase (Get-ConfigurationData)
$autoCsvPath = (Get-ConfigurationData).srcDataLocation + "\" + (Get-ConfigurationData).srcAutoFileName
Import-AutoData (ConnectTo-SQLDatabase (Get-ConfigurationData)) $autoCsvPath