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
