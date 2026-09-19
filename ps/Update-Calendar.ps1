# web/calendar.html に埋め込まれている const DATA / const TODAY を、生成済みのJSONで置き換える。
#   DATA  <- reports/b3l_2026/calendar_data.json (Export-CalendarData.ps1 が作る)
#   TODAY <- reports/today_picks.json            (Get-TodayPicks.ps1 が作る)
# 行そのものを差し替えるだけなので、HTML側の見た目やスクリプトには触らない。
# 使い方: powershell -File ps\Update-Calendar.ps1
param(
    [string]$Html         = "web/calendar.html",
    [string]$CalendarJson = "reports/b3l_2026/calendar_data.json",
    [string]$TodayJson    = "reports/today_picks.json",
    [switch]$NoBackup
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Set-Location (Get-ProjectRoot)

$htmlPath = Resolve-ProjectPath $Html
if (-not (Test-Path $htmlPath)) { throw "$Html がない" }

function Get-CompactJson {
    # 読み込んだJSONをそのまま1行で埋め込む(整形されていても圧縮して入れる)
    param([string]$Path, [string]$Label)
    $full = Resolve-ProjectPath $Path
    if (-not (Test-Path $full)) { Write-Warning "${Label}: $Path がないので据え置く"; return $null }
    $raw = [IO.File]::ReadAllText($full).TrimStart([char]0xFEFF).Trim()
    try { $null = $raw | ConvertFrom-Json } catch { throw "$Path がJSONとして読めない: $_" }
    if ($raw -match "[\r\n]") { $raw = ($raw | ConvertFrom-Json) | ConvertTo-Json -Depth 12 -Compress }
    return $raw
}

$text = [IO.File]::ReadAllText($htmlPath)
$nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
$lines = $text -split [regex]::Escape($nl)

$replaced = @{}
foreach ($t in @(@{ name = "DATA"; json = (Get-CompactJson -Path $CalendarJson -Label "DATA") },
                 @{ name = "TODAY"; json = (Get-CompactJson -Path $TodayJson  -Label "TODAY") })) {
    if ($null -eq $t.json) { continue }
    $hit = $false
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "^(\s*)const $($t.name)\s*=") {
            $lines[$i] = "$($Matches[1])const $($t.name) = $($t.json);"
            $replaced[$t.name] = $i + 1
            $hit = $true
            break
        }
    }
    if (-not $hit) { throw "calendar.html に const $($t.name) の行が見つからない" }
}
if ($replaced.Count -eq 0) { throw "置き換えるJSONが1つもない" }

if (-not $NoBackup) {
    $bak = "$htmlPath.bak"
    [IO.File]::Copy($htmlPath, $bak, $true)
    Write-Host "backup: $Html.bak"
}

[IO.File]::WriteAllText($htmlPath, ($lines -join $nl), (New-Object Text.UTF8Encoding $false))
foreach ($k in $replaced.Keys) { Write-Host ("replaced const {0} (line {1})" -f $k, $replaced[$k]) }
Write-Host "saved $Html ($([Math]::Round((Get-Item $htmlPath).Length / 1KB))KB)"
