# List of PowerShell scripts to accomplish simple tasks

# From initial folder, get list of ".xlsx" files in that folder and all subfolders
gci -recurse -file -include "*.xlsx" | % Fullname > "list_of_files.txt"

# Append text to existing file
gci "file.txt" | ac -value "This is some more text"

# Script to concatenate csv files
$dir = ".\*.csv"
gci $dir | % {
    $variable = "$($_.Name)`n$(Get-content $_.FullName)`n"
    Add-Content -Value $variable -Path .\Output.csv
    }
	
# Script to convert xlsx to csv files

# Read in xlsx files
$dir = ".\*.xlsx"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false; $Excel.DisplayAlerts = $false

# Open each xlsx file and save as ".csv" instead of ".xlsx"
gci $dir | % {
    $wb = $Excel.Workbooks.Open($_)
    $new = (Split-Path $_.FullName -Parent) + "\" + $_.BaseName + ".csv"
    $wb.SaveAs($new,6)
}
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

# Deletes all ".xlsx" files
$dir | % {Remove-Item -path $_}

# Moves all files from parent folder and subdirectories into separate folder, and deletes the subdirectories

gci -Recurse -File | move-item -destination "D:\sepfolder"
gci -recurse -directory | remove-item
