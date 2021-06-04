$dir = ".\*.csv"
gci $dir | % {
    $variable = "$($_.Name)`n$(Get-content $_.FullName)`n"
    Add-Content -Value $variable -Path .\Output.csv
    }