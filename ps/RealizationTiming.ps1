# A label is observable at sell-date opening; forecast is at forecast-date close.
function Get-RealizedHistory {
    param([object[]]$Rows, [string]$AsOf)
    @($Rows | Where-Object { $_.sell_date -and [string]::CompareOrdinal([string]$_.sell_date, $AsOf) -le 0 -and -not [double]::IsNaN([double]$_.rv_pct2) } | Sort-Object sell_date,target_date)
}
function Get-RealizedLags {
    param([object[]]$Rows, [string]$AsOf)
    $h = @(Get-RealizedHistory -Rows $Rows -AsOf $AsOf)
    $v = @([double]::NaN,[double]::NaN,[double]::NaN)
    $ws = @(1,5,22)
    for ($a=0; $a -lt 3; $a++) {
        if ($h.Count -ge $ws[$a]) {
            $sum=0.0
            for ($b=$h.Count-$ws[$a]; $b -lt $h.Count; $b++) { $sum += $h[$b].rv_pct2 }
            $v[$a]=$sum/$ws[$a]
        }
    }
    return ,$v
}
