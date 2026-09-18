# JPXの東証上場銘柄一覧(xlsx)から、指定した市場区分の銘柄を data/raw/universe/*.csv (code,name,sector) に書き出す。
# Excelが無くても読めるよう、xlsx(ZIP圧縮されたXML)を直接解析する。
# 使い方: powershell -File ps\Import-JpxListing.ps1 -XlsxPath data\raw\universe\jpx_listed_202608.xlsx -Market "プライム（内国株式）" -OutCsv data\raw\universe\prime.csv

param(
    [Parameter(Mandatory)][string]$XlsxPath,
    [Parameter(Mandatory)][string]$Market,
    [Parameter(Mandatory)][string]$OutCsv
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryXml {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ColumnIndex {
    # "C12" -> 2 (0始まり)
    param([string]$CellRef)
    $letters = ($CellRef -replace '[0-9]', '')
    $n = 0
    foreach ($ch in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1) }
    return $n - 1
}

$fullPath = $XlsxPath
if (-not [System.IO.Path]::IsPathRooted($fullPath)) { $fullPath = Join-Path (Get-ProjectRoot) $fullPath }
$zip = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
try {
    $shared = New-Object System.Collections.Generic.List[string]
    $ssXml = Read-ZipEntryXml -Zip $zip -Name "xl/sharedStrings.xml"
    foreach ($si in $ssXml.DocumentElement.ChildNodes) {
        # <t>直下の文字列と、書式付きで分割された<r><t>を連結する。ふりがな(<rPh>)は含めない
        $sb = New-Object System.Text.StringBuilder
        foreach ($node in $si.ChildNodes) {
            if ($node.LocalName -eq "t") { [void]$sb.Append($node.InnerText) }
            elseif ($node.LocalName -eq "r") {
                foreach ($rn in $node.ChildNodes) { if ($rn.LocalName -eq "t") { [void]$sb.Append($rn.InnerText) } }
            }
        }
        $shared.Add($sb.ToString())
    }
    $sheet = Read-ZipEntryXml -Zip $zip -Name "xl/worksheets/sheet1.xml"
} finally {
    $zip.Dispose()
}

$rows = New-Object System.Collections.Generic.List[string[]]
foreach ($row in $sheet.worksheet.sheetData.row) {
    $values = New-Object 'string[]' 16
    foreach ($c in $row.c) {
        $col = Get-ColumnIndex -CellRef $c.r
        if ($col -ge 16) { continue }
        $v = [string]$c.v
        if ($c.t -eq "s") { $v = $shared[[int]$v] }
        elseif ($c.t -eq "inlineStr") { $v = [string]$c.is.t }
        $values[$col] = $v
    }
    $rows.Add($values)
}

$header = $rows[0]
$colCode = [Array]::IndexOf($header, "コード")
$colName = [Array]::IndexOf($header, "銘柄名")
$colMarket = [Array]::IndexOf($header, "市場・商品区分")
$colSector = [Array]::IndexOf($header, "33業種区分")
if ($colCode -lt 0 -or $colName -lt 0 -or $colMarket -lt 0 -or $colSector -lt 0) {
    throw "unexpected header: $($header -join ' | ')"
}

$out = New-Object System.Collections.Generic.List[object]
$marketCounts = @{}
for ($i = 1; $i -lt $rows.Count; $i++) {
    $r = $rows[$i]
    $m = $r[$colMarket]
    if (-not $marketCounts.ContainsKey($m)) { $marketCounts[$m] = 0 }
    $marketCounts[$m]++
    if ($m -ne $Market) { continue }
    $out.Add([PSCustomObject]@{ code = $r[$colCode]; name = $r[$colName]; sector = $r[$colSector] })
}

$outFull = Resolve-ProjectPath $OutCsv
$out | Export-Csv -Path $outFull -NoTypeInformation -Encoding UTF8
Write-Host "market counts:"
$marketCounts.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { Write-Host ("  {0,-24} {1,5}" -f $_.Key, $_.Value) }
Write-Host "saved $($out.Count) rows ($Market) to $outFull"
