$txt = [IO.File]::ReadAllText("C:\Users\e.colwell\cbb_tool\data.js")
$D = ($txt.Substring($txt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json
Write-Host "=== teams whose STARTER slot labels are not unique ===" -ForegroundColor Cyan
$bad=0; $tot=0
foreach ($t in $D.teams) {
  $s = @($t.players | Where-Object { $_.kind -eq 'STARTER' })
  $tot++
  $u = @($s | Select-Object -ExpandProperty slot -Unique)
  if ($u.Count -ne $s.Count) {
    $bad++
    if ($bad -le 8) {
      Write-Host ("  {0,-18} slots: {1}   ({2} players -> {3} unique)" -f $t.name, (($s | ForEach-Object { $_.slot }) -join ','), $s.Count, $u.Count)
    }
  }
}
Write-Host ""
Write-Host ("teams with duplicate starter slots: {0} of {1}" -f $bad, $tot) -ForegroundColor Yellow
Write-Host ""
Write-Host "=== worst case: how many players share one slot label ===" -ForegroundColor Cyan
$D.teams | ForEach-Object { $_.players | Group-Object slot | Where-Object { $_.Count -gt 1 } } |
  Group-Object Count | Sort-Object Name | ForEach-Object { Write-Host ("  {0} players share a slot label: {1} occurrences" -f $_.Name, $_.Count) }
