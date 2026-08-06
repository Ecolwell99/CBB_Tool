$txt = [IO.File]::ReadAllText("C:\Users\e.colwell\cbb_tool\data.js")
$D = ($txt.Substring($txt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json
$bad=0
foreach ($t in $D.teams) {
  $ids = @($t.players | ForEach-Object { $_.id })
  $u = @($ids | Select-Object -Unique)
  if ($ids.Count -ne $u.Count) { $bad++; Write-Host ("DUPLICATE ids: {0}" -f $t.name) -ForegroundColor Red }
}
Write-Host ("teams with duplicate player ids: {0} of {1}" -f $bad, $D.teams.Count) -ForegroundColor $(if($bad){'Red'}else{'Green'})
$bc = $D.teams | Where-Object { $_.name -eq 'Boston College' }
Write-Host ""
Write-Host "Boston College (was the worst case: G,G,G,F,F):" -ForegroundColor Cyan
$bc.players | Select-Object -First 7 | ForEach-Object { Write-Host ("  id={0,-3} slot={1,-4} {2}" -f $_.id, $_.slot, $_.name) }
