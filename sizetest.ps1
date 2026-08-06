$hdr = @('Name','Team','Conf','GP','MinPct','ORtg','Usage','eFG','TS','ORBpct','DRBpct','ASTpct','TOpct','FTM','FTA','FTpct','TwoPM','TwoPA','TwoPct','ThreePM','ThreePA','ThreePct','BLKpct','STLpct','FTr')
for ($i=26; $i -le 70; $i++) { $hdr += "c$i" }
$tv = Import-Csv "C:\Users\e.colwell\cbb_tool\torvik_2026.csv" -Header $hdr
function D($x){ $d=0.0; if([double]::TryParse([string]$x,[ref]$d)){return $d}; return 0.0 }

# team games, needed for the MPG fix
$tg=@{}
foreach($p in $tv){ $t=[string]$p.Team; if(-not $t){continue}; $g=D $p.GP; if(-not $tg.ContainsKey($t) -or $g -gt $tg[$t]){$tg[$t]=$g} }

# compact: array-of-arrays, only the fields the model uses
$rows=@()
foreach($p in $tv){
  $gp=D $p.GP; if($gp -le 0){continue}
  $fga=(D $p.TwoPA)+(D $p.ThreePA); if($fga -le 0){continue}
  $t=[string]$p.Team
  $g=if($tg.ContainsKey($t)){$tg[$t]}else{32}
  $mpg=if($gp){[math]::Round(0.4*(D $p.MinPct)*$g/$gp,1)}else{0}
  $rows += ,@($p.Name,$t,$gp,$mpg,
    [math]::Round($fga/$gp,2),
    [math]::Round(((D $p.TwoPM)+(D $p.ThreePM))/$gp,2),
    [math]::Round((D $p.ThreePA)/$gp,2),
    [math]::Round((D $p.ThreePM)/$gp,2))
}
$json = $rows | ConvertTo-Json -Compress -Depth 3
$bytes = [Text.Encoding]::UTF8.GetByteCount($json)
Write-Host ("players kept        : {0}" -f $rows.Count)
Write-Host ("compact index size  : {0:N0} KB" -f ($bytes/1KB))
$cur = (Get-Item "C:\Users\e.colwell\cbb_tool\data.js").Length
Write-Host ("current data.js     : {0:N0} KB" -f ($cur/1KB))
Write-Host ("combined            : {0:N0} KB" -f (($bytes+$cur)/1KB))
Write-Host ("raw torvik CSV      : {0:N0} KB  (for reference)" -f ((Get-Item "C:\Users\e.colwell\cbb_tool\torvik_2026.csv").Length/1KB))
