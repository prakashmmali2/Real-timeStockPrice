# push.ps1 — runs in repo root
git add -A
$dt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
git commit -m "Client save from Excel @ $dt" 2>$null
git push
