$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json
Start-ScheduledTask -TaskName $config.TaskName