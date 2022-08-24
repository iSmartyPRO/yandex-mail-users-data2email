$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

$t = Get-ScheduledTask -TaskName $config.TaskName
$t.Actions