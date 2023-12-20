# Specify the directory containing the files
$directoryPath = "C:\Users\caiop\OneDrive\desktop\output"

# Specify the string to search for and replace in the file names
$searchString = Read-Host -Prompt "Stringa da cercare"
$replaceString = Read-Host -Prompt "Stringa da cambiare"

# Get the list of files in the directory
$files = Get-ChildItem -Path $directoryPath

# Loop through each file and replace the string in the file name
foreach ($file in $files) {
    $oldFileName = $file.Name
    $newFileName = $oldFileName -replace $searchString, $replaceString

    # Check if the string was found in the file name
    if ($oldFileName -ne $newFileName) {
        $newFilePath = Join-Path -Path $directoryPath -ChildPath $newFileName
        $file | Rename-Item -NewName $newFileName -Force
        Write-Output "Renamed file '$oldFileName' to '$newFileName'"
    }
}

#Import excel
Import-Module ImportExcel

#excel file path
$excelFilePath = "C:\Users\atacai\Desktop\testExcel.xlsx"

# Leggo i dati di Excel
$excelData = Import-Excel -Path $excelFilePath

# se serve per controllo, stampa i dati di excel
#$excelData

#Lista di oggetti
$objectList = @()

#ciclo i dati del excel e le salvo in un oggetto, colona/riga
foreach ($row in $excelData) {
	$name = $row.Name
	$id = $row.id
	
	$object1 = [PSCustomObject]@{
    Name = $row.Name
    Id  = $row.id
	}
	
	$objectList += $object1
}

# Variabili per la connesione con SQL server
$serverName = "ATACAI-NB"
$databaseName = "hrport"
$userId = "sa"
$password = "Project1234"

# String di connessione
$connectionString = "Server=$serverName;Database=$databaseName;User Id=$userId;Password=$password;"

# Creo la connessione con il DB
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

try {
    # Apro la connessione
    $connection.Open()
	
	foreach ($obj in $objectList) {
	
    $sqlCommand = $connection.CreateCommand()
	
	#costruisco la string di connessione, con le variabli @
    $sqlCommand.CommandText = "UPDATE utenti SET nome = @Name WHERE idUtente = @Id"
	
	#imposto le variabili per ogni dato del oggetto
	$sqlCommand.Parameters.AddWithValue("@Name", $obj.Name) | Out-Null
    $sqlCommand.Parameters.AddWithValue("@Id", $obj.Id) | Out-Null
	
    $result = $sqlCommand.ExecuteNonQuery()
	}

	Write-Host "Rows affected: $result"
	
} catch {
    Write-Host "Error updating data: $_"
} finally {
    # Close the connection
    $connection.Close()
}

#Install-Module -Name SqlServer -Force -AllowClobber
#Set-ExecutionPolicy RemoteSigned
#Install-Module -Name ImportExcel -Force -AllowClobber

