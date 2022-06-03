$d = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
$y = $d.Split('-')[0]
$t = Get-Date -Format "yyyy-MM-dd"
$dirSource = (Get-Location).ToString() + '\0000-00'
$dirTarget = (Get-Location).ToString() + '\' + $y + '\' + $d
Copy-Item -Path $dirSource -Destination $dirTarget -Recurse
Set-Location $dirTarget
Get-ChildItem -Recurse -Include "*.xlsx" | Rename-Item -NewName { $_.Name -replace '\[DATA\]', $d }
Get-ChildItem -Recurse -Include "*.xlsx" | Rename-Item -NewName { $_.Name -replace '\[TODAY\]', $t }
explorer $dirTarget
exit
