$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json
Unregister-ScheduledTask -TaskName $config.TaskName