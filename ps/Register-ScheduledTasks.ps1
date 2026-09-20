# 日次更新(17時)と取得の確認(翌朝8時)を Windows のタスクスケジューラに登録する。
# 管理者権限は不要。ログオンしている間だけ動く(パスワードを預けなくて済む)。
# 登録:   powershell -File ps\Register-ScheduledTasks.ps1
# 確認:   powershell -File ps\Register-ScheduledTasks.ps1 -Show
# 取り消し: powershell -File ps\Register-ScheduledTasks.ps1 -Unregister
param(
    [string]$UpdateTime = "17:00",
    [string]$CheckTime  = "08:00",
    [int]$Budget = 500000,
    [switch]$Unregister,
    [switch]$Show
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\Common.ps1"
$root = Get-ProjectRoot

$tasks = @(
    @{ name = "kabu-b2 daily update"; script = "Invoke-DailyUpdate.ps1"; time = $UpdateTime; desc = "営業日の大引け後に株価を取り直し、カレンダーと翌営業日の銘柄を更新する" }
    @{ name = "kabu-b2 daily check";  script = "Test-DailyUpdate.ps1";  time = $CheckTime;  desc = "前営業日の17時の更新が通ったかを確かめ、駄目なら取り直す" }
)

if ($Show) {
    foreach ($t in $tasks) {
        $task = Get-ScheduledTask -TaskName $t.name -ErrorAction SilentlyContinue
        if (-not $task) { Write-Host ("{0,-22} 未登録" -f $t.name); continue }
        $info = Get-ScheduledTaskInfo -TaskName $t.name
        Write-Host ("{0,-22} {1,-8} 次回 {2}  前回 {3} (結果 {4})" -f $t.name, $task.State, $task.Triggers[0].StartBoundary.Substring(11, 5), $info.LastRunTime, $info.LastTaskResult)
    }
    exit 0
}

if ($Unregister) {
    foreach ($t in $tasks) {
        if (Get-ScheduledTask -TaskName $t.name -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $t.name -Confirm:$false
            Write-Host "削除: $($t.name)"
        } else { Write-Host "なし: $($t.name)" }
    }
    exit 0
}

foreach ($t in $tasks) {
    $taskArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$root\ps\$($t.script)`" -Budget $Budget"
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $taskArgs -WorkingDirectory $root
    # 毎日動かし、休業日かどうかはスクリプト側で判定する(祝日を扱うため)
    $trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::ParseExact($t.time, "HH:mm", $null))
    $settings = New-ScheduledTaskSettingsSet `
        -StartWhenAvailable `
        -DontStopIfGoingOnBatteries `
        -AllowStartIfOnBatteries `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Hours 3)
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

    if (Get-ScheduledTask -TaskName $t.name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $t.name -Confirm:$false
    }
    Register-ScheduledTask -TaskName $t.name -Description $t.desc `
        -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
    Write-Host ("登録: {0,-22} 毎日 {1}  -> ps\{2}" -f $t.name, $t.time, $t.script)
}

Write-Host ""
Write-Host "PCがスリープ・電源断のときは動かない。StartWhenAvailable を入れてあるので、"
Write-Host "起動後に取りこぼした回をできるだけ早く実行する。"
Write-Host "状態の確認: powershell -File ps\Register-ScheduledTasks.ps1 -Show"
