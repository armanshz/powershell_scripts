

for ($i = 1; $i -le 10; $i++) {
    $dir = ".\file$i.txt"; $string = "This is file $i"
    Add-Content -path $dir -value $string
}
$dir = "E:\python" ;cd $dir; $arr = gc "$dir\*.csv"; Write-Host $arr.Length


# $arr = $a.Split(" "); Write-Host $arr[0]
# $str = "hello12723619now23132"; $arr = $str -split "\D+"; Write-Host $arr