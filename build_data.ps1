<#
  CBB Lineup Tool - data builder
  Reads the roster xlsx READ-ONLY (copies to temp; the source file is never written to),
  pulls Barttorvik season stats, joins them, emits data.js for index.html

  Usage:
    powershell -ExecutionPolicy Bypass -File build_data.ps1
    powershell -ExecutionPolicy Bypass -File build_data.ps1 -Year 2027
    powershell -ExecutionPolicy Bypass -File build_data.ps1 -SkipDownload
#>
param(
  [string]$Xlsx,
  [int]$Year = 2026,
  [string]$OutDir = "C:\Users\e.colwell\cbb_tool",
  [switch]$SkipDownload
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---------- 0. find the roster sheet ----------
# Not hardcoded to one filename: re-downloading gives "... (1).xlsx", and the sheet may
# live in Downloads, the tool folder, or OneDrive.  Newest matching file wins.
if (-not $Xlsx) {
  $searchDirs = @(
    $OutDir,
    (Join-Path $env:USERPROFILE 'Downloads'),
    (Join-Path $env:USERPROFILE 'Desktop'),
    (Join-Path $env:USERPROFILE 'Documents')
  )
  if ($env:OneDrive) { $searchDirs += @($env:OneDrive, (Join-Path $env:OneDrive 'Downloads'), (Join-Path $env:OneDrive 'Documents')) }
  if ($env:OneDriveCommercial) { $searchDirs += $env:OneDriveCommercial }

  $found = @()
  foreach ($d in ($searchDirs | Select-Object -Unique)) {
    if (-not $d -or -not (Test-Path -LiteralPath $d)) { continue }
    $found += Get-ChildItem -LiteralPath $d -Filter '*.xlsx' -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -match 'roster' -and $_.Name -notlike '~$*' }
  }
  if ($found.Count -eq 0) {
    Write-Host ""
    Write-Host "Could not find a roster .xlsx automatically." -ForegroundColor Red
    Write-Host "Looked in:" -ForegroundColor DarkGray
    ($searchDirs | Select-Object -Unique) | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    Write-Host ""
    Write-Host "Fix: either drop the sheet in $OutDir, or pass the path:" -ForegroundColor Yellow
    Write-Host '  .\build_data.ps1 -Xlsx "C:\path\to\sheet.xlsx"' -ForegroundColor Yellow
    throw "No roster sheet found."
  }
  $pick = $found | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  $Xlsx = $pick.FullName
  if ($found.Count -gt 1) {
    Write-Host ("Found {0} roster files - using the most recently modified:" -f $found.Count) -ForegroundColor Yellow
    $found | Sort-Object LastWriteTime -Descending | ForEach-Object {
      $mark = if ($_.FullName -eq $Xlsx) { '  ->' } else { '    ' }
      Write-Host ("{0} {1}  ({2})" -f $mark, $_.FullName, $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) -ForegroundColor DarkGray
    }
  }
}
$xlsxInfo = Get-Item -LiteralPath $Xlsx
Write-Host ("Sheet: {0}" -f $xlsxInfo.FullName) -ForegroundColor Cyan
Write-Host ("Saved: {0}  ({1:N0} KB)" -f $xlsxInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), ($xlsxInfo.Length/1KB)) -ForegroundColor DarkGray
$ageHrs = ((Get-Date) - $xlsxInfo.LastWriteTime).TotalHours
if ($ageHrs -gt 168) {
  Write-Host ("NOTE: this file was last saved {0:N0} days ago - is it the current version?" -f ($ageHrs/24)) -ForegroundColor Yellow
}

$tmp = Join-Path ([IO.Path]::GetTempPath()) "cbb_build"
if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }

# ---------- 1. unzip xlsx (source is only ever read) ----------
if (-not (Test-Path $Xlsx)) { throw "Roster sheet not found: $Xlsx" }
$work = Join-Path $tmp "book.zip"
Copy-Item -LiteralPath $Xlsx -Destination $work
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::ExtractToDirectory($work, (Join-Path $tmp "x"))
$xl = Join-Path $tmp "x\xl"
Write-Host "Read roster sheet (source untouched)" -ForegroundColor Green

function ReadXml($path) {
  $sr = New-Object System.IO.StreamReader($path, [System.Text.Encoding]::UTF8)
  $t = $sr.ReadToEnd(); $sr.Close()
  return [xml]$t
}

# ---------- 2. shared strings ----------
$strings = New-Object System.Collections.ArrayList
$ssPath = Join-Path $xl "sharedStrings.xml"
if (Test-Path $ssPath) {
  $ss = ReadXml $ssPath
  foreach ($si in $ss.sst.si) {
    $txt = ''
    if ($si.t -ne $null) { $txt = if ($si.t -is [string]) { $si.t } else { $si.t.'#text' } }
    elseif ($si.r -ne $null) { $txt = ($si.r | ForEach-Object { if ($_.t -is [string]) { $_.t } else { $_.t.'#text' } }) -join '' }
    [void]$strings.Add([string]$txt)
  }
}

# ---------- 3. sheet order / names ----------
$wb = ReadXml (Join-Path $xl "workbook.xml")
$rels = ReadXml (Join-Path $xl "_rels\workbook.xml.rels")
$relMap = @{}
foreach ($r in $rels.Relationships.Relationship) {
  $relMap[$r.Id] = ($r.Target -replace '^/?xl/', '' -replace '^worksheets/', '')
}
$sheetOrder = @()
foreach ($s in $wb.workbook.sheets.sheet) {
  $rid = $s.id
  if (-not $rid) { $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') }
  $file = $relMap[$rid]
  if ($file) { $sheetOrder += [pscustomobject]@{ Name = $s.name; File = $file } }
}

# ---------- 4. parse rosters ----------
function CellVal($cell, $strings) {
  $v = $cell.v
  if ($cell.t -eq 's' -and $v -ne $null) { return [string]$strings[[int]$v] }
  if ($cell.t -eq 'inlineStr') { return [string]$cell.is.t }
  if ($v -eq $null) { return $null }
  return [string]$v
}
$KEYWORDS = @('coach:', 'team notes', 'key departures', 'notes', 'height', 'player notes')
function IsKeyword($s) {
  if (-not $s) { return $false }
  $l = $s.ToLower().Trim()
  foreach ($k in $KEYWORDS) { if ($l.StartsWith($k)) { return $true } }
  return $false
}

$roster = New-Object System.Collections.ArrayList
foreach ($sh in $sheetOrder) {
  $p = Join-Path $xl "worksheets\$($sh.File)"
  if (-not (Test-Path $p)) { continue }
  $doc = ReadXml $p
  $team = ''
  $section = ''
  foreach ($row in $doc.worksheet.sheetData.row) {
    $c = @{}
    foreach ($cell in $row.c) {
      $ref = $cell.r -replace '[0-9]', ''
      $v = CellVal $cell $strings
      if ($v -ne $null -and $v.Trim() -ne '') { $c[$ref] = $v.Trim() }
    }
    if ($c.Count -eq 0) { continue }
    $A = $c['A']; $B = $c['B']

    # keyword rows checked FIRST, so "Key Departures:" is never read as a team header
    if (IsKeyword $A) {
      $la = $A.ToLower()
      if ($la.StartsWith('key departures')) { $section = 'DEPARTED' }
      elseif ($la.StartsWith('team notes')) { $section = 'NOTES' }
      continue
    }
    # a lone value in column A is a team name
    if ($A -and -not $B -and $c.Count -eq 1) { $team = $A; $section = 'ROSTER'; continue }
    if ($section -ne 'ROSTER') { continue }
    if (-not $A -or -not $B) { continue }
    if ($A.Length -gt 6) { continue }

    $kind = if ($A -match '^B\d+$') { 'BENCH' } else { 'STARTER' }
    [void]$roster.Add([pscustomobject]@{
        Conf = $sh.Name; Team = $team; Slot = $A; Kind = $kind; Name = $B; Year = $c['D']
        PPG = $c['E']; RPG = $c['F']; APG = $c['G']; SPG = $c['H']; BPG = $c['I']
        FGpct = $c['J']; TPpct = $c['K']; TPAg = $c['L']; Notes = $c['M']
      })
  }
}
Write-Host ("Parsed {0} players / {1} teams" -f $roster.Count, (($roster | Select-Object -ExpandProperty Team -Unique).Count)) -ForegroundColor Green

# ---------- 5. Barttorvik stats ----------
$csv = Join-Path $OutDir "torvik_$Year.csv"
if (-not ($SkipDownload -and (Test-Path $csv))) {
  $url = "https://barttorvik.com/getadvstats.php?year=$Year&csv=1"
  Write-Host "Downloading $url ..." -NoNewline
  try {
    Invoke-WebRequest -Uri $url -OutFile $csv -UseBasicParsing -TimeoutSec 90
    Write-Host " ok" -ForegroundColor Green
  } catch {
    if (Test-Path $csv) { Write-Host " FAILED - using cached copy" -ForegroundColor Yellow }
    else { throw "Torvik download failed and no cached copy exists: $_" }
  }
}
$hdr = @('Name', 'Team', 'Conf', 'GP', 'MinPct', 'ORtg', 'Usage', 'eFG', 'TS', 'ORBpct', 'DRBpct',
  'ASTpct', 'TOpct', 'FTM', 'FTA', 'FTpct', 'TwoPM', 'TwoPA', 'TwoPct', 'ThreePM', 'ThreePA',
  'ThreePct', 'BLKpct', 'STLpct', 'FTr')
for ($i = 26; $i -le 70; $i++) { $hdr += "c$i" }
$tv = Import-Csv $csv -Header $hdr
Write-Host ("Torvik rows: {0}" -f $tv.Count) -ForegroundColor Green

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
$idx = @{}
foreach ($p in $tv) {
  $k = Norm $p.Name
  if (-not $k) { continue }
  if (-not $idx.ContainsKey($k)) { $idx[$k] = New-Object System.Collections.ArrayList }
  [void]$idx[$k].Add($p)
}
function D($x) {
  $d = 0.0
  if ([double]::TryParse([string]$x, [ref]$d)) { return $d }
  return 0.0
}
function TorvikPPG($t) {
  $gp = D $t.GP
  if ($gp -le 0) { return 0.0 }
  return (2 * (D $t.TwoPM) + 3 * (D $t.ThreePM) + (D $t.FTM)) / $gp
}

# Torvik's Min% is a share of the TEAM's whole season, not of the games this player
# appeared in.  So total minutes = 0.4 * Min% * teamGames -- dividing by the player's
# own GP would inflate the per-minute rates of anyone who missed time (a 6-game
# player came out ~5x too high before this).  Team games = max GP on that roster.
$teamGames = @{}
foreach ($p in $tv) {
  $tm = [string]$p.Team
  if (-not $tm) { continue }
  $gp = D $p.GP
  if (-not $teamGames.ContainsKey($tm) -or $gp -gt $teamGames[$tm]) { $teamGames[$tm] = $gp }
}
function TeamG($t) {
  $tm = [string]$t.Team
  if ($tm -and $teamGames.ContainsKey($tm) -and $teamGames[$tm] -gt 0) { return $teamGames[$tm] }
  return 32.0
}

# ---------- 6. join ----------
$stats = [ordered]@{ torvik = 0; sheet = 0; none = 0; rejected = 0; ambig = 0 }
$teams = @{}
$auditRows = New-Object System.Collections.ArrayList

foreach ($p in $roster) {
  if (-not $teams.ContainsKey($p.Team)) {
    $teams[$p.Team] = [pscustomobject]@{ name = $p.Team; conf = $p.Conf; players = (New-Object System.Collections.ArrayList) }
  }

  $sheetPPG = $null
  if ($p.PPG) { $sheetPPG = D $p.PPG }

  $priorSchool = $null
  if ($p.Notes -and $p.Notes -match '^([A-Za-z][A-Za-z\.\s&\-]{1,28}?)\s+transfer') { $priorSchool = $Matches[1].Trim() }

  $k = Norm $p.Name
  $cands = @()
  if ($idx.ContainsKey($k)) { $cands = @($idx[$k]) }

  $pick = $null
  $why = ''
  if ($cands.Count -eq 1) {
    $pick = $cands[0]; $why = 'name'
  } elseif ($cands.Count -gt 1) {
    $stats.ambig++
    # 1) confirm with the sheet's own PPG - the sheet is built from the same season
    if ($sheetPPG -ne $null) {
      $best = $cands | Sort-Object { [math]::Abs((TorvikPPG $_) - $sheetPPG) } | Select-Object -First 1
      if ([math]::Abs((TorvikPPG $best) - $sheetPPG) -lt 0.35) { $pick = $best; $why = 'name+ppg' }
    }
    # 2) prior school parsed out of the Notes column
    if (-not $pick -and $priorSchool) {
      $ps = ($priorSchool -replace '[^A-Za-z]', '').ToLower()
      $m = @($cands | Where-Object { $_.Team -and ($_.Team -replace '[^A-Za-z]', '').ToLower().StartsWith($ps) })
      if ($m.Count -gt 0) { $pick = $m[0]; $why = 'name+school' }
    }
    # 3) give up and take the higher-minutes player, flagged
    if (-not $pick) {
      $pick = $cands | Sort-Object { - ((D $_.MinPct) * (D $_.GP)) } | Select-Object -First 1
      $why = 'name+minutes'
    }
  }

  # sanity gate: sheet has stats but the matched player scores nothing like them -> wrong person
  $warn = $null
  if ($pick -and $sheetPPG -ne $null) {
    $diff = [math]::Abs((TorvikPPG $pick) - $sheetPPG)
    if ($diff -gt 4.0) {
      $warn = ("name matched {0} but PPG {1:N1} vs sheet {2:N1} - match rejected" -f $pick.Team, (TorvikPPG $pick), $sheetPPG)
      $pick = $null
      $stats.rejected++
    } elseif ($diff -gt 1.5) {
      $warn = ("PPG differs from sheet by {0:N1}" -f $diff)
    }
  }
  if ($pick -and $why -eq 'name+minutes') {
    $warn = (@($warn, 'ambiguous name - could not confirm which player') | Where-Object { $_ }) -join '; '
  }

  # id must be unique within the team -- slot is a POSITION label (G, F, G/F) and is
  # routinely shared by 2-3 players, so it cannot be used to identify a row.
  $o = [ordered]@{ id = $teams[$p.Team].players.Count; slot = $p.Slot; name = $p.Name; year = $p.Year; kind = $p.Kind; notes = $p.Notes; sheetPPG = $sheetPPG }

  if ($pick) {
    $gp = D $pick.GP
    $fga = (D $pick.TwoPA) + (D $pick.ThreePA)
    $fgm = (D $pick.TwoPM) + (D $pick.ThreePM)
    $tg = TeamG $pick
    $totMin = 0.4 * (D $pick.MinPct) * $tg
    if ($gp -gt 0 -and $fga -gt 0) {
      $o.src = 'torvik'; $o.why = $why; $o.gp = $gp
      $o.teamGP = $tg
      # minutes per game the player actually appeared in
      $o.mpg = [math]::Round($totMin / $gp, 1)
      $o.fga = [math]::Round($fga / $gp, 3)
      $o.fgm = [math]::Round($fgm / $gp, 3)
      $o.tpa = [math]::Round((D $pick.ThreePA) / $gp, 3)
      $o.tpm = [math]::Round((D $pick.ThreePM) / $gp, 3)
      $o.fta = [math]::Round((D $pick.FTA) / $gp, 3)
      $o.ftm = [math]::Round((D $pick.FTM) / $gp, 3)
      $o.tvTeam = $pick.Team
      $o.usage = [math]::Round((D $pick.Usage), 1)
      $stats.torvik++
    } else { $pick = $null }
  }

  if (-not $pick) {
    if ($sheetPPG -ne $null -and (D $p.FGpct) -gt 0) {
      # fallback: back FGA out of PPG.  FGA = (PPG - 3PM - FT) / (2*FG%), FT assumed 20% of PPG.
      # Less accurate than Torvik (mean 2.7pts of 3PA-frequency error) - flagged in the UI.
      $fgp = (D $p.FGpct) / 100
      $tpp = (D $p.TPpct) / 100
      $tpa = D $p.TPAg
      $tpm = $tpa * $tpp
      $fga = ($sheetPPG - $tpm - 0.20 * $sheetPPG) / (2 * $fgp)
      $floored = $false
      if ($fga -lt $tpa) { $fga = $tpa; $floored = $true }
      $o.src = 'sheet'; $o.why = 'reconstructed from sheet PPG/FG%/3P%'
      $o.gp = $null; $o.mpg = $null
      $o.fga = [math]::Round($fga, 3)
      $o.fgm = [math]::Round($fga * $fgp, 3)
      $o.tpa = [math]::Round($tpa, 3)
      $o.tpm = [math]::Round($tpm, 3)
      $o.fta = $null
      $o.ftm = [math]::Round(0.20 * $sheetPPG, 3)
      if ($floored) {
        $warn = (@($warn, 'FGA floored at 3PA - free-throw estimate too high for this shot profile') | Where-Object { $_ }) -join '; '
      }
      $stats.sheet++
    } else {
      $o.src = 'none'; $o.why = 'no college stats on file'
      foreach ($f in 'gp', 'mpg', 'fga', 'fgm', 'tpa', 'tpm', 'fta', 'ftm') { $o[$f] = $null }
      $stats.none++
    }
  }
  $o.warn = $warn
  [void]$teams[$p.Team].players.Add([pscustomobject]$o)

  if ($warn) {
    [void]$auditRows.Add([pscustomobject]@{ Team = $p.Team; Slot = $p.Slot; Name = $p.Name; Src = $o.src; Issue = $warn })
  }
}

# ---------- 6b. diff against the previous build ----------
# So a refresh reports what actually changed in the sheet rather than silently swapping data.
$prevPath = Join-Path $OutDir "data.js"
$diffLines = @()
if (Test-Path -LiteralPath $prevPath) {
  try {
    $ptxt = [IO.File]::ReadAllText($prevPath)
    $prev = ($ptxt.Substring($ptxt.IndexOf('{')).TrimEnd(';')) | ConvertFrom-Json
    $prevMap = @{}
    foreach ($t in $prev.teams) {
      foreach ($p in $t.players) { $prevMap["$($t.name)|$($p.name)"] = $p }
    }
    $curMap = @{}
    foreach ($tn in $teams.Keys) {
      foreach ($p in $teams[$tn].players) { $curMap["$tn|$($p.name)"] = $p }
    }
    $added = @($curMap.Keys | Where-Object { -not $prevMap.ContainsKey($_) })
    $removed = @($prevMap.Keys | Where-Object { -not $curMap.ContainsKey($_) })
    $slotChg = @()
    foreach ($k in $curMap.Keys) {
      if ($prevMap.ContainsKey($k)) {
        $a = $prevMap[$k]; $b = $curMap[$k]
        if ("$($a.slot)" -ne "$($b.slot)") { $slotChg += ("{0}: {1} -> {2}" -f $k.Replace('|', ' / '), $a.slot, $b.slot) }
      }
    }
    if ($added.Count) { $diffLines += "  + {0} player(s) added" -f $added.Count }
    if ($removed.Count) { $diffLines += "  - {0} player(s) removed" -f $removed.Count }
    if ($slotChg.Count) { $diffLines += "  ~ {0} depth-chart change(s)" -f $slotChg.Count }
    $script:diffDetail = @{ added = $added; removed = $removed; slot = $slotChg }
  } catch { $diffLines += "  (could not diff previous build)" }
}

# ---------- 7. emit ----------
$order = @{}
$i = 0
foreach ($s in $sheetOrder) { $order[$s.Name] = $i; $i++ }
$teamList = @($teams.Values | Sort-Object @{e = { $order[$_.conf] } }, name)
$confs = @($sheetOrder | ForEach-Object { $_.Name } | Where-Object { $n = $_; ($teamList | Where-Object { $_.conf -eq $n }).Count -gt 0 })

$payload = [ordered]@{
  generated  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
  sheetSaved = $xlsxInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
  statsYear  = "{0}-{1}" -f ($Year - 1), $Year.ToString().Substring(2)
  torvikYear = $Year
  source     = "Depth chart: {0}  |  Stats: barttorvik.com (year={1})" -f (Split-Path $Xlsx -Leaf), $Year
  counts     = $stats
  conferences = $confs
  teams      = $teamList
}
$json = $payload | ConvertTo-Json -Depth 8 -Compress
$enc = New-Object System.Text.UTF8Encoding($false)
# data.js (not .json) so index.html works straight off the filesystem without a web server
[IO.File]::WriteAllText((Join-Path $OutDir "data.js"), "window.CBB_DATA = $json;", $enc)

# ---------- 7b. full Torvik index ----------
# Every D1 player, so an uploaded roster can be joined in the browser even for players
# who were not in the last build.  A page opened from file:// cannot call barttorvik.com,
# so the whole stats universe has to be on disk.  Array-of-arrays keeps it ~360KB.
$idxRows = New-Object System.Collections.ArrayList
foreach ($p in $tv) {
  $gp = D $p.GP
  if ($gp -le 0) { continue }
  $fga = (D $p.TwoPA) + (D $p.ThreePA)
  if ($fga -le 0) { continue }
  $tgm = TeamG $p
  $mpg = [math]::Round(0.4 * (D $p.MinPct) * $tgm / $gp, 1)
  # ppg is carried so the browser can confirm identity the same way this script does --
  # it is the strongest signal for telling apart two players with the same name.
  $ppg = [math]::Round((2 * (D $p.TwoPM) + 3 * (D $p.ThreePM) + (D $p.FTM)) / $gp, 2)
  [void]$idxRows.Add(@(
      [string]$p.Name,
      [string]$p.Team,
      $gp,
      $mpg,
      [math]::Round($fga / $gp, 2),
      [math]::Round(((D $p.TwoPM) + (D $p.ThreePM)) / $gp, 2),
      [math]::Round((D $p.ThreePA) / $gp, 2),
      [math]::Round((D $p.ThreePM) / $gp, 2),
      $ppg
    ))
}
$idxPayload = [ordered]@{
  built  = (Get-Date -Format 'yyyy-MM-dd HH:mm')
  year   = $Year
  fields = @('name', 'team', 'gp', 'mpg', 'fga', 'fgm', 'tpa', 'tpm', 'ppg')
  rows   = $idxRows
}
$idxJson = $idxPayload | ConvertTo-Json -Depth 4 -Compress
[IO.File]::WriteAllText((Join-Path $OutDir "index_torvik.js"), "window.CBB_INDEX = $idxJson;", $enc)
Write-Host ("Torvik index: {0} players -> index_torvik.js ({1:N0} KB)" -f $idxRows.Count, ([Text.Encoding]::UTF8.GetByteCount($idxJson) / 1KB)) -ForegroundColor Green
if ($auditRows.Count -gt 0) {
  $auditRows | Export-Csv -NoTypeInformation -Encoding UTF8 (Join-Path $OutDir "audit.csv")
}

Write-Host ""
Write-Host "=== BUILD COMPLETE ===" -ForegroundColor Cyan
Write-Host ("teams              : {0}" -f $teamList.Count)
Write-Host ("torvik stats       : {0}" -f $stats.torvik) -ForegroundColor Green
Write-Host ("sheet reconstructed: {0}" -f $stats.sheet) -ForegroundColor Yellow
Write-Host ("no data            : {0}" -f $stats.none) -ForegroundColor DarkGray
Write-Host ("ambiguous names    : {0} resolved   bad matches rejected: {1}" -f $stats.ambig, $stats.rejected)

if ($diffLines.Count -gt 0) {
  Write-Host ""
  Write-Host "Changes vs previous build:" -ForegroundColor Cyan
  $diffLines | ForEach-Object { Write-Host $_ -ForegroundColor White }
  $dd = $script:diffDetail
  if ($dd) {
    $show = 6
    if ($dd.added.Count) {
      Write-Host "  added:" -ForegroundColor DarkGray
      $dd.added | Select-Object -First $show | ForEach-Object { Write-Host ("    + {0}" -f $_.Replace('|', ' / ')) -ForegroundColor DarkGray }
      if ($dd.added.Count -gt $show) { Write-Host ("    ... and {0} more" -f ($dd.added.Count - $show)) -ForegroundColor DarkGray }
    }
    if ($dd.removed.Count) {
      Write-Host "  removed:" -ForegroundColor DarkGray
      $dd.removed | Select-Object -First $show | ForEach-Object { Write-Host ("    - {0}" -f $_.Replace('|', ' / ')) -ForegroundColor DarkGray }
      if ($dd.removed.Count -gt $show) { Write-Host ("    ... and {0} more" -f ($dd.removed.Count - $show)) -ForegroundColor DarkGray }
    }
    if ($dd.slot.Count) {
      Write-Host "  depth chart:" -ForegroundColor DarkGray
      $dd.slot | Select-Object -First $show | ForEach-Object { Write-Host ("    ~ {0}" -f $_) -ForegroundColor DarkGray }
      if ($dd.slot.Count -gt $show) { Write-Host ("    ... and {0} more" -f ($dd.slot.Count - $show)) -ForegroundColor DarkGray }
    }
  }
} elseif (Test-Path -LiteralPath $prevPath) {
  Write-Host ""
  Write-Host "No roster changes vs previous build." -ForegroundColor DarkGray
}

Write-Host ""
Write-Host ("-> {0}\data.js" -f $OutDir)
if ($auditRows.Count -gt 0) { Write-Host ("-> {0}\audit.csv  ({1} rows needing a look)" -f $OutDir, $auditRows.Count) -ForegroundColor Yellow }
Write-Host ""
Write-Host "Now refresh the tool in your browser (Ctrl+F5)." -ForegroundColor Green
