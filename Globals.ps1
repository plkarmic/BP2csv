#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ConfigurationData
{
	$configData = (Get-Content .\config.json) | ConvertFrom-Json
	
	return $configData
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
		$connection,
		[Parameter(Mandatory = $true, Position = 2)]
		$csvPath
	)
	
	Get-Content $csvPath >> $PWD\temp.csv
	$row = Import-Csv $PWD\temp.csv -Delimiter ";"
	$row
}

#ConnectTo-SQLDatabase (Get-ConfigurationData)
$autoCsvPath = (Get-ConfigurationData).srcDataLocation + "\" + (Get-ConfigurationData).srcAutoFileName
Import-AutoData (ConnectTo-SQLDatabase (Get-ConfigurationData)) $autoCsvPath