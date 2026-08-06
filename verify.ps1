# Independent check of the numbers index.html should be showing.
# Recomputes 3PA frequency / Make% / P(3|made) straight from data.js
# using the same per-40 logic, so the JS can be compared against it.
$ErrorActionPreference = 'Stop'
$OutDir = "C:\Users\e.colwell\cbb_tool"
$txt = [IO.File]::ReadAllText((Join-Path $OutDir "data.js"))
$json = $txt.Substring($txt.IndexOf('{'))
$json = $json.TrimEnd(';')
$D = $json | ConvertFrom-Json

$MIN_FLOOR = 8
function Per40($p) {
  if ($p.fga -eq $null) { return $null }
  $m = $p.mpg
  $k = 1.0
  if ($m -ne $null -and $m -gt 0) { $k = 40.0 / [math]::Max([double]$m, $MIN_FLOOR) }
  return [pscustomobject]@{ fga = [double]$p.fga * $k; fgm = [double]$p.fgm * $k; tpa = [double]$p.tpa * $k; tpm = [double]$p.tpm * $k }
}
function Calc($players) {
  $fga = 0.0; $fgm = 0.0; $tpa = 0.0; $tpm = 0.0; $n = 0
  foreach ($p in $players) {
    $r = Per40 $p
    if (-not $r) { continue }
    $fga += $r.fga; $fgm += $r.fgm; $tpa += $r.tpa; $tpm += $r.tpm; $n++
  }
  if ($fga -le 0) { return $null }
  return [pscustomobject]@{ N = $n; FGA = $fga; Freq3 = $tpa / $fga; Make = $fgm / $fga; P3 = $(if ($fgm) { $tpm / $fgm } else { 0 }) }
}

Write-Host "=== data.js sanity ===" -ForegroundColor Cyan
Write-Host ("teams: {0}   generated: {1}   stats: {2}" -f $D.teams.Count, $D.generated, $D.statsYear)

$all = @()
foreach ($t in $D.teams) {
  $five = @($t.players | Where-Object { $_.kind -eq 'STARTER' -and $_.fga -ne $null } | Select-Object -First 5)
  if ($five.Count -lt 5) { continue }
  $r = Calc $five
  if (-not $r) { continue }
  $torv = @($five | Where-Object { $_.src -eq 'torvik' }).Count
  $all += [pscustomobject]@{ Team = $t.name; Conf = $t.conf; FGA40 = [math]::Round($r.FGA, 1)
    Freq3 = [math]::Round(100 * $r.Freq3, 1); Make = [math]::Round(100 * $r.Make, 1); P3 = [math]::Round(100 * $r.P3, 1); Torvik5 = $torv }
}
Write-Host ""
Write-Host ("complete projected fives: {0}" -f $all.Count) -ForegroundColor Green
$a = $all | Measure-Object FGA40 -Average -Minimum -Maximum
$f = $all | Measure-Object Freq3 -Average -Minimum -Maximum
$m = $all | Measure-Object Make -Average -Minimum -Maximum
$p = $all | Measure-Object P3 -Average -Minimum -Maximum
Write-Host ""
Write-Host "=== league-wide, projected starting fives (per-40 normalized) ===" -ForegroundColor Cyan
Write-Host ("SumFGA/40  avg={0,5:N1}  min={1,5:N1}  max={2,5:N1}   <- plausible band ~48-80" -f $a.Average, $a.Minimum, $a.Maximum)
Write-Host ("3PA freq   avg={0,5:N1}% min={1,5:N1}% max={2,5:N1}%   <- D1 ~38-39%" -f $f.Average, $f.Minimum, $f.Maximum)
Write-Host ("Make%      avg={0,5:N1}% min={1,5:N1}% max={2,5:N1}%   <- D1 ~45%" -f $m.Average, $m.Minimum, $m.Maximum)
Write-Host ("3 of makes avg={0,5:N1}% min={1,5:N1}% max={2,5:N1}%" -f $p.Average, $p.Minimum, $p.Maximum)

Write-Host ""
Write-Host "=== outliers worth eyeballing ===" -ForegroundColor Yellow
$all | Where-Object { $_.FGA40 -lt 48 -or $_.FGA40 -gt 82 } | Sort-Object FGA40 | Format-Table -AutoSize | Out-String -Width 140 | Write-Host

Write-Host "=== sample teams (compare these to the browser) ===" -ForegroundColor Cyan
$all | Sort-Object Team | Select-Object -First 15 | Format-Table -AutoSize | Out-String -Width 140 | Write-Host

Write-Host "=== highest / lowest 3PA frequency ===" -ForegroundColor Cyan
$all | Sort-Object Freq3 -Descending | Select-Object -First 6 | Format-Table Team, Conf, Freq3, Make, P3 -AutoSize | Out-String -Width 140 | Write-Host
$all | Sort-Object Freq3 | Select-Object -First 6 | Format-Table Team, Conf, Freq3, Make, P3 -AutoSize | Out-String -Width 140 | Write-Host

# swap-delta demo: pull the single biggest available swap for one team
Write-Host "=== swap-delta example ===" -ForegroundColor Cyan
$t = $D.teams | Where-Object { $_.name -eq 'Arizona' } | Select-Object -First 1
if (-not $t) { $t = $D.teams | Where-Object { (@($_.players | Where-Object { $_.kind -eq 'STARTER' -and $_.fga -ne $null }).Count -ge 5) } | Select-Object -First 1 }
$five = @($t.players | Where-Object { $_.kind -eq 'STARTER' -and $_.fga -ne $null } | Select-Object -First 5)
$base = Calc $five
Write-Host ("{0} projected five: 3PA freq {1:N1}%  Make {2:N1}%  3-of-makes {3:N1}%" -f $t.name, (100 * $base.Freq3), (100 * $base.Make), (100 * $base.P3))
$bench = @($t.players | Where-Object { $_.kind -eq 'BENCH' -and $_.fga -ne $null })
$best = @()
foreach ($i in $bench) {
  foreach ($o in $five) {
    $lineup = @($five | Where-Object { $_.slot -ne $o.slot }) + @($i)
    $r = Calc $lineup
    if ($r) { $best += [pscustomobject]@{ In = $i.name; Out = $o.name; dFreq = [math]::Round(100 * ($r.Freq3 - $base.Freq3), 1); dMake = [math]::Round(100 * ($r.Make - $base.Make), 1) } }
  }
}
$best | Sort-Object { - [math]::Abs($_.dFreq) } | Select-Object -First 8 | Format-Table -AutoSize | Out-String -Width 140 | Write-Host
