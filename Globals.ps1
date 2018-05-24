#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ConfigurationData ()
{
    Param
    (
        [parameter(Mandatory = $true, Position = 1)]
		[string]$inputFile
    )
    $configData = (Get-Content $inputFile) | ConvertFrom-Json
	
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

function Get-AutoID()
{
    Param
    (
        [Parameter(Mandatory = $true,Position = 1)]
        $SQLconnection,
        [Parameter(Mandatory = $true,Position = 2)]
        $nrRej
    )
    $sqlQuery= "SELECT ID FROM AUTO WHERE [NR_Rejestracyjny] = '{0}'" -f $nrRej

        
    $sqlCmd = New-Object System.Data.OleDb.OleDbCommand $sqlQuery, $SQLconnection
    $sqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $sqlCmd 
    $dataSet = New-Object System.Data.DataSet
    [void]$sqlAdapter.Fill($dataSet)
    
    return $dataSet.Tables[0].ID
}

function Get-FakturaID()
{
    Param
    (
        [Parameter(Mandatory = $true,Position = 1)]
        $SQLconnection,
        [Parameter(Mandatory = $true,Position = 2)]
        $faktura
    )
    $sqlQuery= "SELECT ID FROM Faktura WHERE [NR_Faktury] = '{0}'" -f $faktura

        
    $sqlCmd = New-Object System.Data.OleDb.OleDbCommand $sqlQuery, $SQLconnection
    $sqlAdapter = New-Object System.Data.OleDb.OleDbDataAdapter $sqlCmd 
    $dataSet = New-Object System.Data.DataSet
    [void]$sqlAdapter.Fill($dataSet)
    
    return $dataSet.Tables[0].ID
}

function Import-Data-Transaction ()
{
	param (
		[Parameter(Mandatory = $true, Position = 1)]
		$SQLConnection,
		[Parameter(Mandatory = $true, Position = 2)]
		$csvPath
	)

	$rows = Get-Content $csvPath
    $i = 0
    while($i -lt 6)
    {   
        $rows[$i] >> $pwd\temp-headers2.txt
        $i++
    }

    $header = @{}
    $rowsHeader = Get-Content $pwd\temp-headers2.txt
    rm $pwd\temp-headers2.txt -Force
    foreach ($row in $rowsHeader)
    {
        if(($row.Split(",")).Count -gt 2)
        {
            $key = ($row.Split(",")[0]) -replace (":","") -replace ('"','')
            $r = 1
            $value = ""
            while ($r -lt ($row.Split(",")).Count)
            {
                $value += $row.Split(",")[$r] -replace (":","") -replace ('"','')
                $r++
            }
            #$hash = $row.Split(",")[1] + $row.Split(",")[2]
            $h =  @{$key=$value}
        }
        else 
        {
            $h =  @{($row.Replace('"','').Replace(":","")).Split(",")[0]=($row.Replace('"','').Replace(":","")).Split(",")[1]}
        }
        $header += $h
    }
    $csvHead = "Data transakcji, Numer transakcj, Osrodek Kosztów, Nr wystawcy karty, Numer Klienta, Numer seryjny karty, Uzytkownik, Numer rejestracyjny, Stan licznika, Nr faktury ICC, Transakcja w dzien roboczy., Stacja, Nazwa stacji, Numer produktu, Nazwa produktu, Ilosc, Cena jednostkowa, Wartosc netto (faktura), Kwota VAT (faktura), Wartosc Brutto  (faktura), Wartosc Brutto , Stawka VAT, Kod kraju, Waluta, Cena jednostkowa2, Wartosc Netto, Kwota VAT, Liczba Oplat, Unikalny numer transakcji., Wskaznik RC, Oryginalna wartosc netto., Oryginalna wartosc VAT., Ogyginalna wartosc brutto., Oryginalna stawka VAT., Oryginalny kod kraju., Oryginalna waluta."
    $csvHead > $pwd\temp-out.csv
    $rows = Get-Content $csvPath
    $rows.Replace('"',"") | Select-Object -Skip 8 >> $pwd\temp-out.csv
    $rows = Import-Csv $pwd\temp-out.csv
    rm $pwd\temp-out.csv
    #$rows

    #Import to FAKTURA from $header
    $sqlCmd = New-Object System.Data.OleDb.OleDbCommand
    $transaction = $SqlConnection.BeginTransaction("Insert into Faktura and Transakcja")
    $transactionError = 0
    $sqlCmd.Transaction = $transaction
   
    $sqlQuery = "INSERT INTO [dbo].[FAKTURA] ([NR_FAKTURY],[DATA],[SUM_TRANS],[SUM_NETT],[SUM_BRUTT],[NR_NOT_OBCI],[WALUTA]) VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}')" `
        -f $header.Faktura, $header.Data, $header.Suma, $header.'Suma Netto', $header.Brutto, $header.'Numer Noty Obciazeniowej', $header.Waluta
    try
    {
        $sqlCmd.CommandText = $sqlQuery
        $sqlCmd.Connection = $SQLConnection
        if(($sqlCmd.ExecuteNonQuery()) -eq 1)
        {
            Log-ToFile "Object $($header['Faktura']) added into Faktura database"
        }
    } 
    catch
    {
        if ($_.Exception.ErrorRecord.Exception.Message -like "*Cannot insert duplicate key in object*")
        {
            Log-ToFile "Object $($header['Faktura']) already exists"
        }
        else 
        {
            Log-ToFile "Object $($header['Faktura']) not added, unknown error"    
        }
        $transactionError = 1
    }   
    
    #Import to Transakcje from $rows
    foreach($row in $rows)
    {
        
        $autoID = Get-AutoID $SQLConnection  $row.'Numer rejestracyjny'
        $fakturaID = Get-FakturaID $SQLConnection $header.Faktura
        
        $sqlQuery = "INSERT INTO [dbo].[TRANSAKCJA] `
        ([AUTO_ID],
        [ID_FAKTURA],
        [DATA_TRANS],
        [NR_TRANS],
        [OSRODEK_KOSZT],
        [NR_WYST_KARTY],
        [NR_KLIENTA],
        [NR_SER_KARTY],
        [UZYTKOWNIK],
        [NR_REJ],
        [STAN_LICZ],
        [NR_FAK_ICC],
        [TRANS_DZIEN_ROB],
        [STACJA],
        [NAZWA_STACJI],
        [NR_PROD],
        [NAZWA_PROD],
        [ILOSC],
        [FAK_CENA_JEDN],
        [FAK_WARTOSC_NET],
        [FAK_KWOTA_VAT],
        [FAK_WARTOSC_BRUT],
        [WARTOSC_BRUT],
        [VAT],
        [KOD_KRAJ],
        [WALUTA],
        [CEN_JEDN],
        [WARTOSC_NET],
        [KWOTA_VAT],
        [LICZBA_OPLAT],
        [UNIK_NR_TRANS],
        [WSK_RC],
        [ORG_WART_NET],
        [ORG_WAR_VAT],
        [ORG_WART_BRUT],
        [ORG_STAW_VAT],
        [ORG_KOD_KRAJ],
        [ORG_WALUTA]
        )
        VALUES
        (
            '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','{36}','{37}'
        )"`
        -f 
        $autoID,`
        $fakturaID,`
        $row.'Data transakcji',`
        $row.'Numer transakcj',`
        $row.'Osrodek Kosztów',`
        $row.'Nr wystawcy karty',`
        $row.'Numer Klienta',`
        $row.'Numer seryjny karty',`
        $row.'Uzytkownik',`
        $row.'Numer rejestracyjny',`
        $row.'Stan licznika',`
        $row.'Nr faktury ICC',`
        $row.'Transakcja w dzien roboczy.',`
        $row.'Stacja',`
        $row.'Nazwa stacji',`
        $row.'Numer produktu',`
        $row.'Nazwa produktu ',`
        $row.'Ilosc',`
        $row.'Cena jednostkowa',`
        $row.'Wartosc netto (faktura)',`
        $row.'Kwota VAT (faktura)',`
        $row.'Wartosc Brutto  (faktura)',`
        $row.'Wartosc Brutto',`
        $row.'Stawka VAT',`
        $row.'Kod kraju',`
        $row.'Waluta',`
        $row.'Cena jednostkowa2',`
        $row.'Wartosc Netto',`
        $row.'Kwota VAT',`
        $row.'Liczba Oplat',`
        $row.'Unikalny numer transakcji.',`
        $row.'Wskaznik RC',`
        $row.'Oryginalna wartosc netto.',`
        $row.'Oryginalna wartosc VAT.',`
        $row.'Ogyginalna wartosc brutto.',`
        $row.'Oryginalna stawka VAT.',`
        $row.'Oryginalny kod kraju. ',`
        $row.'Oryginalna waluta.'
        
        Log-ToFile "SQL query: " + $sqlQuery

        try
        {
            $sqlCmd.CommandText = $sqlQuery
            $sqlCmd.Connection = $SQLConnection
            if(($sqlCmd.ExecuteNonQuery()) -eq 1)
            {
                Log-ToFile "Object $($row.'Numer transakcj') ,added into Faktura database"
            }
        } 
        catch
        {
            if ($_.Exception.ErrorRecord.Exception.Message -like "*Cannot insert duplicate key in object*")
            {
                Log-ToFile "Object $($row.'Numer transakcj') already exists"
            }
            else 
            {
                Log-ToFile "Object $($row.'Numer transakcj') not added, unknown error"    
            }
            $transactionError += 1
        }   
    }
    
    if($transactionError -gt 0)
    {
        Log-ToFile "Insert error, transaction $($transaction) rolled back" 
        $transaction.Rollback()
    }
    else 
    {
        Log-ToFile "Insert OK, transaction $($transaction) commited" 
        $transaction.Commit()
    }

}

$configData = Get-ConfigurationData \\wassv006\INSTALL\Projects\BP\BP-sourceCode\BP2csv\config.json
#ConnectTo-SQLDatabase ($configData)
#$autoCsvPath = $configData.srcDataLocation + "\" + $configData.srcAutoFileName
#Import-AutoData (ConnectTo-SQLDatabase $configData) $autoCsvPath

Import-Data-Transaction (ConnectTo-SQLDatabase $configData) "\\wassv006\Install\Projects\BP\Input data\BPPL42244-1810057218.csv"