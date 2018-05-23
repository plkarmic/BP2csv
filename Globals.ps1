#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ConfigurationData
{
	$configData = (Get-Content .\configuration.json) | ConvertFrom-Json
	
	return $configData
}

function ConnectTo-SQLDatabase ()
{
	
	param (
		# Parameter help description
		[Parameter(Mandatory = $true, Position = 1)]
		$configuration,
		[Parameter(Mandatory = $true, Position = 2)]
		$databaseName
	)
	
	$dataSource = $configuration.databaseServer
	$database = $databaseName
	
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


