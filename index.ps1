$config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

Import-Module ".\libs\iScript.psm1"

$all_users = Get-YaEmailList

$active_users = $all_users | ? {$_.isEnabled -eq $true} | select id, lastName, firstName, nickname, email, position, gender, createdAt, updatedAt
$inactive_users = $all_users | ? {$_.isEnabled -eq $false} | select id, lastName, firstName, nickname, email, position, gender, createdAt, updatedAt

$active_users_html = ConvertTo-HtmlTable -header ('ID', 'Фамилия', 'Имя', 'Электронная почта', 'Должность', 'Пол', 'Дата создания', 'Дата обновления') -dataKeys ('id', 'lastName', 'firstName', 'email', 'position', 'gender', 'createdAt', 'updatedAt') -dataValues $active_users
$inactive_users_html = ConvertTo-HtmlTable -header ('ID', 'Фамилия', 'Имя', 'Электронная почта', 'Должность', 'Пол', 'Дата создания', 'Дата обновления') -dataKeys ('id', 'lastName', 'firstName', 'email', 'position', 'gender', 'createdAt', 'updatedAt') -dataValues $inactive_users

$template = Get-Content ".\templates\yaEmails.html" -Encoding "UTF8"

$emailBody = $template.Replace('{{$active_users}}', $active_users_html)
$emailBody = $emailBody.Replace('{{$inactive_users}}', $inactive_users_html)

Push-iMail  -SMTPServer $config.SMTPServer `
            -SMTPPort $config.Port `
            -SMTPUser $config.SMTPUser `
            -SMTPPass $config.SMTPPass `
            -Priority $config.Priority `
            -EnableSsl $config.EnableSsl `
            -Sender $config.Sender `
            -ReplyTo $config.ReplyTo `
            -From $config.From `
            -To $config.To `
            -Subject $config.Subject `
            -Body $emailBody