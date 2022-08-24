$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json


$uri = "https://oauth.yandex.ru/authorize?response_type=token&client_id=$($config.clientId)"

(iwr -URI $uri -Method GET -SkipHeaderValidation).Links