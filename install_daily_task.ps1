param(
    [string]$TaskName = "Whizzard Vehicle Open Tickets Daily",
    [string]$DailyAt = "08:00"
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPath = Join-Path $ProjectDir "daily_open_ticket_report.py"
$PythonExe = (Get-Command python.exe -ErrorAction Stop).Source

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Daily report script not found: $ScriptPath"
}

$time = [datetime]::ParseExact($DailyAt, "HH:mm", [Globalization.CultureInfo]::InvariantCulture)
$action = New-ScheduledTaskAction -Execute $PythonExe -Argument ('"{0}"' -f $ScriptPath) -WorkingDirectory $ProjectDir
$trigger = New-ScheduledTaskTrigger -Daily -At $time
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Downloads the five current-month Whizzard ticket CSVs and creates the open-ticket Excel report." -Force | Out-Null
Write-Host "Scheduled task '$TaskName' installed for $DailyAt every day."
