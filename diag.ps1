$ErrorActionPreference = 'Stop'
$txt = [IO.File]::ReadAllText("C:\Users\e.colwell\cbb_tool\data.js")
$D = ($txt.Substring($txt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json

foreach ($nm in 'USC', 'San Diego St', 'Northwestern', 'Indiana', 'Louisville') {
  $t = $D.teams | Where-Object { $_.name -eq $nm } | Select-Object -First 1
  if (-not $t) { continue }
  Write-Host ""
  Write-Host "=== $nm projected five ===" -ForegroundColor Cyan
  $five = @($t.players | Where-Object { $_.kind -eq 'STARTER' -and $_.fga -ne $null } | Select-Object -First 5)
  $five | ForEach-Object {
    $k = 1.0
    if ($_.mpg -ne $null -and $_.mpg -gt 0) { $k = 40.0 / [math]::Max([double]$_.mpg, 8) }
    Write-Host ("  {0,-22} src={1,-6} GP={2,-4} MPG={3,-5} FGA/g={4,-6:N2} -> FGA/40={5,-6:N1} (x{6:N2})  tvTeam={7}" -f `
        $_.name, $_.src, $_.gp, $_.mpg, [double]$_.fga, ([double]$_.fga * $k), $k, $_.tvTeam)
  }
}

Write-Host ""
Write-Host "=== distribution of MPG among players in projected fives ===" -ForegroundColor Cyan
$mp = @()
foreach ($t in $D.teams) {
  foreach ($p in @($t.players | Where-Object { $_.kind -eq 'STARTER' -and $_.fga -ne $null } | Select-Object -First 5)) {
    if ($p.mpg -ne $null) { $mp += [double]$p.mpg }
  }
}
$mp = $mp | Sort-Object
Write-Host ("n={0}  min={1:N1}  p5={2:N1}  p25={3:N1}  median={4:N1}  mean={5:N1}  max={6:N1}" -f `
    $mp.Count, $mp[0], $mp[[int]($mp.Count * 0.05)], $mp[[int]($mp.Count * 0.25)], $mp[[int]($mp.Count * 0.5)], ($mp | Measure-Object -Average).Average, $mp[-1])
Write-Host ("players under 15 mpg in a projected five: {0}" -f (@($mp | Where-Object { $_ -lt 15 }).Count))
Write-Host ("players under 20 mpg in a projected five: {0}" -f (@($mp | Where-Object { $_ -lt 20 }).Count))

Write-Host ""
Write-Host "=== FGA/40 distribution for individual players (Torvik rows) ===" -ForegroundColor Cyan
$f40 = @()
foreach ($t in $D.teams) {
  foreach ($p in $t.players) {
    if ($p.src -eq 'torvik' -and $p.mpg -ne $null -and $p.mpg -gt 0) {
      $f40 += [double]$p.fga * 40.0 / [math]::Max([double]$p.mpg, 8)
    }
  }
}
$f40 = $f40 | Sort-Object
Write-Host ("n={0}  median={1:N1}  p90={2:N1}  p99={3:N1}  max={4:N1}   <- a real high-usage player is ~18-22 FGA/40" -f `
    $f40.Count, $f40[[int]($f40.Count * 0.5)], $f40[[int]($f40.Count * 0.9)], $f40[[int]($f40.Count * 0.99)], $f40[-1])
Write-Host ("players over 25 FGA/40 (implausible): {0}" -f (@($f40 | Where-Object { $_ -gt 25 }).Count))
