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
		$connection,
		[Parameter(Mandatory = $true, Position = 2)]
		$csvPath
	)

<<<<<<< HEAD
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
=======
	$row = Import-Csv $csvPath -Delimiter ","
	$row
>>>>>>> 6d954e14f2816400b455218bc5397750a97ebe1a
}

function Add-UsersIntoSQLDatabase()
{
    param(
        [Parameter(Mandatory=$true, Position=1)]
        $SQLconnection,
        [Parameter(Mandatory=$true, Position=2)]
        $ADusers
    )

   
        
        
        
        $sqlQuery = "INSERT INTO [dbo].[UsersAccount] ([samAccountName],[Name],[LastName],[DistinguishedName],[BusinessUnit],[Company],[OfficeCode],[Type],[DisplayName],[SID],[Created],[Status], [Modified])`
        VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}')" `
        -f $user.samAccountName, $user.GivenName, $user.SurName, $user.DistinguishedName, $user.extensionAttribute1, $user.extensionAttribute1, $user.extensionAttribute15,`
            $user.extensionAttribute6, $user.DisplayName, $user.SID, $user.Created, $user.Enabled, $user.Modified

        

        #$sqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $sqlCmd = New-Object System.Data.OleDb.OleDbCommand
        try{
            $sqlCmd.CommandText = $sqlQuery
            $sqlCmd.Connection = $SQLConnection
            if(($sqlCmd.ExecuteNonQuery()) -eq 1){
                Log-ToFile "Object $($user.samAccountName) added into database"
            }
        }catch{
            if ($_.Exception.ErrorRecord.Exception.Message -like "*Cannot insert duplicate key in object 'dbo.UsersAccount'*"){
                Log-ToFile "Object $($user.samAccountName) already exists in database"

                $modifiedDB = $dataSet.Tables[0].Modified
                $modifiedAD = (Get-ADUser $($user.samAccountName) -Properties Modified).Modified

                if((get-date($modifiedDB)) -lt (get-date($modifiedAD)))
                {
                    Log-ToFile "AD object: $user.samAccountName has been modified, database update required..."
                    Update-SQLDatabase $SQLConnection $user "USR" 
                }

                #Sprawdzic date ostatniej modyfikacji w AD oraz w bazie => jeżeli starsza w bazie to zauktualizować rekord
            }
        }
        
        # WORKAROUND: have to be duplicated for first run (no accont exist in useraccoutns db) moved to line 77
        $sqlQuery = "SELECT * FROM [dbo].[UsersAccount] WHERE [SID] = '{0}'" -f $user.SID
        $sqlQuery
        $sqlCmd = New-Object System.Data.OleDb.OleDbCommand $sqlQuery, $SQLconnection
        $sqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $sqlCmd 
        $dataSet = New-Object System.Data.DataSet
        [void]$sqlAdapter.Fill($dataSet)

        $dataSet

        $accountID = $dataSet.Tables[0].Id
        $reportPeriod = get-date -Format yyyy.MM

        #FOR TEST PURPISES ONLY

        #$reportPeriod = Get-Date 
        #$reportPeriod=$date.AddMonths(1)
        #$reportPeriod = Get-Date($date) -Format yyyy.MM

        ###

        $hash = Get-MD5Hash ($accountID.ToString() + ($reportPeriod).ToString())

        $sqlQuery = "INSERT INTO [dbo].[UsersAccountHistory] ([AccountID],[Date],[Status],[ReportPeriod], [Hash]) VALUES ('{0}','{1}','{2}','{3}','{4}')" `
                -f $accountID, (Get-Date -Format "yyyy-MM-dd"), $user.Enabled, ($reportPeriod.ToString()), $hash
        $sqlCmd = New-Object System.Data.OleDb.OleDbCommand
        try{
            $sqlCmd.CommandText = $sqlQuery
            $sqlCmd.Connection = $SQLConnection
            if(($sqlCmd.ExecuteNonQuery()) -eq 1){
                Log-ToFile "History information for $($user.samAccountName) object updated"
            }    
        }catch{
            if ($_.Exception.ErrorRecord.Exception.Message -like "*Cannot insert duplicate key in object 'dbo.UsersAccountHistory'*"){
                Log-ToFile "History information for $($user.samAccountName) already exists"
            }
        }

    }
    
    #$sqlConnection.Close()

}


#ConnectTo-SQLDatabase (Get-ConfigurationData)
$autoCsvPath = (Get-ConfigurationData).srcDataLocation + "\" + (Get-ConfigurationData).srcAutoFileName
Import-AutoData (ConnectTo-SQLDatabase (Get-ConfigurationData)) $autoCsvPath