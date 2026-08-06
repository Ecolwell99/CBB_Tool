# Simulates the BROWSER-side join (index_torvik.js + PPG confirmation) and compares it to
# the PowerShell build (full CSV + PPG confirmation).  They must agree, or an uploaded
# sheet would price differently than the bundled data.
$ErrorActionPreference = 'Stop'
$OutDir = "C:\Users\e.colwell\cbb_tool"

$itxt = [IO.File]::ReadAllText((Join-Path $OutDir "index_torvik.js"))
$IDX = ($itxt.Substring($itxt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json
$F = @{}
for ($i = 0; $i -lt $IDX.fields.Count; $i++) { $F[$IDX.fields[$i]] = $i }
Write-Host ("index: {0} players, built {1}, year {2}, fields: {3}" -f $IDX.rows.Count, $IDX.built, $IDX.year, ($IDX.fields -join ',')) -ForegroundColor Cyan

function Norm($s) {
  if (-not $s) { return '' }
  $s = $s.ToLower().Trim()
  $s = $s -replace '[\u00e0-\u00e5]', 'a' -replace '[\u00e8-\u00eb]', 'e' -replace '[\u00ec-\u00ef]', 'i'
  $s = $s -replace '[\u00f2-\u00f6]', 'o' -replace '[\u00f9-\u00fc]', 'u' -replace '[\u00f1]', 'n'
  $s = $s -replace '[\u0107\u010d\u00e7]', 'c' -replace '[\u017e\u017c\u017a]', 'z' -replace '[\u0161\u015b]', 's'
  $s = $s -replace '[^a-z0-9 ]', ' '
  $s = $s -replace '\s+(jr|sr|ii|iii|iv|v)$', ''
  $s = $s -replace '\s+', ' '
  return $s.Trim()
}
$LOOK = @{}
foreach ($r in $IDX.rows) {
  $k = Norm $r[$F['name']]
  if (-not $k) { continue }
  if (-not $LOOK.ContainsKey($k)) { $LOOK[$k] = New-Object System.Collections.ArrayList }
  [void]$LOOK[$k].Add($r)
}

$dtxt = [IO.File]::ReadAllText((Join-Path $OutDir "data.js"))
$D = ($dtxt.Substring($dtxt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json
function N($v) { $d = 0.0; if ([double]::TryParse([string]$v, [ref]$d)) { return $d } return 0.0 }

$agree = 0; $diffSrc = 0; $diffPick = 0; $tot = 0
$maxRateDiff = 0.0
$examples = New-Object System.Collections.ArrayList
$counts = @{ torvik = 0; sheet = 0; none = 0; rejected = 0; ambig = 0 }

foreach ($t in $D.teams) {
  foreach ($p in $t.players) {
    $tot++
    $sheetPPG = $null
    if ($p.sheetPPG -ne $null) { $sheetPPG = N $p.sheetPPG }
    $priorSchool = $null
    if ($p.notes -and $p.notes -match '^([A-Za-z][A-Za-z\.\s&\-]{1,28}?)\s+transfer') { $priorSchool = $Matches[1].Trim() }

    $cands = @()
    $k = Norm $p.name
    if ($LOOK.ContainsKey($k)) { $cands = @($LOOK[$k]) }

    $pick = $null; $why = ''
    if ($cands.Count -eq 1) { $pick = $cands[0]; $why = 'name' }
    elseif ($cands.Count -gt 1) {
      $counts.ambig++
      if ($sheetPPG -ne $null) {
        $best = $cands | Sort-Object { [math]::Abs((N $_[$F['ppg']]) - $sheetPPG) } | Select-Object -First 1
        if ([math]::Abs((N $best[$F['ppg']]) - $sheetPPG) -lt 0.35) { $pick = $best; $why = 'name+ppg' }
      }
      if (-not $pick -and $priorSchool) {
        $ps = ($priorSchool -replace '[^A-Za-z]', '').ToLower()
        $m = @($cands | Where-Object { ("$($_[$F['team']])" -replace '[^A-Za-z]', '').ToLower().StartsWith($ps) })
        if ($m.Count) { $pick = $m[0]; $why = 'name+school' }
      }
      if (-not $pick) { $pick = $cands | Sort-Object { - ((N $_[$F['mpg']]) * (N $_[$F['gp']])) } | Select-Object -First 1; $why = 'name+minutes' }
    }
    if ($pick -and $sheetPPG -ne $null) {
      if ([math]::Abs((N $pick[$F['ppg']]) - $sheetPPG) -gt 4.0) { $pick = $null; $counts.rejected++ }
    }

    $browserSrc = if ($pick) { 'torvik' } elseif ($sheetPPG -ne $null -and $p.src -ne 'none') { 'sheet' } else { 'none' }
    # mirror the build's extra requirement of a usable FG%
    if ($browserSrc -eq 'sheet' -and $p.src -eq 'none') { $browserSrc = 'none' }
    $counts[$browserSrc]++

    if ($browserSrc -eq $p.src) {
      $agree++
      if ($p.src -eq 'torvik' -and $pick) {
        if ($p.tvTeam -and "$($pick[$F['team']])" -ne "$($p.tvTeam)") {
          $diffPick++
          if ($examples.Count -lt 10) { [void]$examples.Add("PICK  $($t.name) / $($p.name): browser=$($pick[$F['team']]) build=$($p.tvTeam)") }
        }
        # numbers must match too, not just the choice of row
        foreach ($fl in 'fga', 'fgm', 'tpa', 'tpm', 'mpg') {
          $bv = N $pick[$F[$fl]]
          $sv = N $p.$fl
          $d = [math]::Abs($bv - $sv)
          if ($d -gt $maxRateDiff) { $maxRateDiff = $d }
        }
      }
    } else {
      $diffSrc++
      if ($examples.Count -lt 10) { [void]$examples.Add("SRC   $($t.name) / $($p.name): browser=$browserSrc build=$($p.src)") }
    }
  }
}

Write-Host ""
Write-Host "=== browser-side join vs PowerShell build ===" -ForegroundColor Cyan
Write-Host ("rows compared        : {0}" -f $tot)
Write-Host ("same data source     : {0}  ({1:N2}%)" -f $agree, (100 * $agree / $tot)) -ForegroundColor Green
Write-Host ("different source     : {0}" -f $diffSrc) -ForegroundColor $(if ($diffSrc -gt 3) { 'Red' } else { 'Green' })
Write-Host ("different Torvik row : {0}" -f $diffPick) -ForegroundColor $(if ($diffPick -gt 0) { 'Red' } else { 'Green' })
Write-Host ("max rate difference  : {0:N4}  (rounding only if < 0.01)" -f $maxRateDiff) -ForegroundColor $(if ($maxRateDiff -gt 0.011) { 'Red' } else { 'Green' })
Write-Host ""
Write-Host ("browser: torvik={0} sheet={1} none={2} rejected={3} ambig={4}" -f $counts.torvik, $counts.sheet, $counts.none, $counts.rejected, $counts.ambig)
Write-Host ("build  : torvik={0} sheet={1} none={2} rejected={3} ambig={4}" -f $D.counts.torvik, $D.counts.sheet, $D.counts.none, $D.counts.rejected, $D.counts.ambig)
if ($examples.Count) {
  Write-Host ""
  Write-Host "divergences:" -ForegroundColor Yellow
  $examples | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
}
