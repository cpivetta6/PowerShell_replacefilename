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

# Read data from Excel
$excelData = Import-Excel -Path $excelFilePath

# Display the data
#$excelData

$objectList = @()

foreach ($row in $excelData) {
	$name = $row.Name
	$age = $row.Age
	
	$object1 = [PSCustomObject]@{
    Name = $row.Name
    Age  = $row.Age
	}
	
	$objectList += $object1
	
	 #Write-Host "$($name) $($age)"
	
	##if ($row.Name -ne "" -or $row.Age -ne "") {
     ##   Write-Host "$($row.Name), $($row.Age)"
    ##}
}

# Display only values without headers
foreach ($obj in $objectList) {
    Write-Host "$($obj.Name) $($obj.Age)"
}


# Variables for SQL Server connection
$serverName = "ATACAI-NB"
$databaseName = "hrport"
$userId = "sa"
$password = "Project1234"

# Construct the connection string
$connectionString = "Server=$serverName;Database=$databaseName;User Id=$userId;Password=$password;"

# Create a SQL connection
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

try {
    # Open the connection
    $connection.Open()

   
	
	foreach ($obj in $objectList) {
	# SQL commands go here
    $sqlCommand = $connection.CreateCommand()
    $sqlCommand.CommandText = "UPDATE utenti SET nome = $obj.Name WHERE idUtente = $obj.Age"
    #$result = $sqlCommand.ExecuteReader()
    #Write-Host "$($obj.Name) $($obj.Age)"
	
	}

	
	<#
    # Process the results or perform other tasks
	 while ($result.Read()) {
        # Access each column by its name or index
        $column1Value = $result["idUtente"]
        $column2Value = $result["email"]

        # Print or process the values
        Write-Host "Column1: $column1Value, Column2: $column2Value"
    }
	#>

} catch {
    Write-Host "Error: $_"
} finally {
    # Close the connection
    $connection.Close()
}

