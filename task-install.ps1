$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

$Action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-NonInteractive -NoLogo -NoProfile -File `"$($config.AppPath)\index.ps1`"" -WorkingDirectory $config.AppPath
$Trigger = New-ScheduledTaskTrigger -Daily -At $config.TaskLaunchTime
$Settings = New-ScheduledTaskSettingsSet -DontStopOnIdleEnd -RestartInterval (New-TimeSpan -Minutes 1) -RestartCount 10 -StartWhenAvailable

$Task = New-ScheduledTask -Action $Action -Trigger $Trigger -Settings $Settings
$Task | Register-ScheduledTask -TaskName $config.TaskName -User "SYSTEM"