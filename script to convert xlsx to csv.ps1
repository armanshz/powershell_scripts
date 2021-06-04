# read in xlsx files

$dir = ".\*.xlsx"
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false; $Excel.DisplayAlerts = $false

# open each xlsx file and save as ".csv" instead of ".xlsx"
gci $dir | % {
    $wb = $Excel.Workbooks.Open($_)
    $new = (Split-Path $_.FullName -Parent) + "\" + $_.BaseName + ".csv"
    $wb.SaveAs($new,6)
}
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

# deletes all ".xlsx" files
$dir | % {Remove-Item -path $_}