# define Error handling
# note: do not change these values
$global:ErrorActionPreference = "Stop"
if($verbose){ $global:VerbosePreference = "Continue" }


#==========================================================================
# FUNCTION Push-iMail
#==========================================================================
Function Push-iMail {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [String]$SMTPServer,
        [String]$SMTPPort = 25,
        [String]$SMTPUser,
        [String]$SMTPPass,
        [String]$Sender,
        [String]$From,
        [String]$ReplyTo = "",
        [String]$Priority = "Normal",
        [Boolean]$EnableSsl = $false,
        [Boolean]$IsBodyHTML = $true,
        [Array]$To,
        [String]$Subject,
        [Array]$Body

    )

    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
          $Smtp = New-Object System.Net.Mail.SMTPClient($SMTPServer,$SMTPPort)
          $Smtp.EnableSsl = $EnableSsl
          $smtp.UseDefaultCredentials = $false;
          $Smtp.Timeout = 30000
          $Smtp.Credentials = New-Object System.Net.NetworkCredential($SMTPUser, $SMTPPass)

          # Create the message
          $mail = New-Object System.Net.Mail.Mailmessage
          $mail.Sender = "$Sender <$From>"
          $mail.From = $From
          $mail.Priority = $Priority
          $mail.ReplyTo = $ReplyTo
          ForEach($recipient in $To) {
            $mail.To.Add($recipient)
          }
          $mail.Subject = $Subject
          $mail.Body = $Body
          $mail.IsBodyHTML=$IsBodyHTML

          # Send send the Mail
          $result = $smtp.send($mail)
          return $result
          #}
            Exit 0
        } catch {
            Write-Host "Error description:"
            Write-Host $_
        }
    }
    end {
        Write-Host "Finished"
    }
}

#==========================================================================
# FUNCTION Get-iLocalUsers
#==========================================================================
Function Get-iLocalUsers {
    <#
        .SYNOPSIS
        Get Local Users with needed info
        .DESCRIPTION
        Get Local Users with needed info
    #>
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            return Get-LocalUser | ? {$_.Enabled -eq $true } | Select Name, FullName, Description, PasswordExpires, UserMayChangePassword, PasswordRequired, PasswordLastSet, LastLogon
        } catch {
            Write-Host "Error description:"
            Exit 1
        }
    }
    end {

    }
}

#==========================================================================
# FUNCTION ConvertTo-ParsedUsers
#==========================================================================
Function ConvertTo-ParsedUsers {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [Object]$users
    )
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            $parsed = @()
            ForEach($user in $users) {
                <# $user #>
                $desc = $user.Description
                if($desc -gt 0) {
                    $descSplited = $desc.Split(";")
                    if($descSplited.Length -eq 1) {
                        $department = $descSplited[0]
                    } else {
                        $position, $TabelNumber = ""
                    }
                    if($descSplited.Length -eq 2) {
                        $department = $descSplited[0]
                        $position = $descSplited[1]
                    } else {
                        $TabelNumber = ""
                    }
                    if($descSplited.Length -eq 3) {
                        $department = $descSplited[0]
                        $position = $descSplited[1]
                        $TabelNumber = $descSplited[2]
                    }
                } else {
                    $department, $position, $TabelNumber = ""
                }
                $parsed += [pscustomobject] @{
                    UserName = $user.Name
                    FullName = $user.FullName
                    Department = $department
                    Position = $position
                    TabelNumber = $TabelNumber
                }
            }
            return $parsed
        } catch {
            Write-Host $_
        }
    }
    end {

    }
}

#==========================================================================
# FUNCTION ConvertTo-ParsedUsers
#==========================================================================
Function Get-ADSecurityGroups {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [Object]$SearchBase
    )
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            $sgGroups = Get-ADGroup -Filter * -Properties description -SearchBase $SearchBase | ? {$_.Name -like "sg_*" } | select Name, description | sort-object name


            <# Get rows for table #>
            $parsed = @()
            ForEach($g in $sgGroups) {
                $description = $g.description
                if($description.Length -ne 0) {
                    $descSplited = $description.Split(";")
                    if($descSplited.Length -eq 1) {
                        $path = $descSplited[0]
                    } else {
                        $accessType = ""
                    }
                    if($descSplited.Length -eq 2) {
                        $path = $descSplited[0]
                        $accessType = $descSplited[1]
                    }
                } else {
                    $path, $accessType = ""
                }



                $adMembers = Get-ADGroupMember -Identity $g.Name
                $members = ""
                if(($adMembers | measure).Count -ne 0) {
                $members += "`n`r`t`t`t<ul>`n`r"
                    ForEach($m in $adMembers) {
                        $members += "`t`t`t`t<li>$($m.name)<br><i>($($m.SamAccountName))</i>;</li>`n`r"
                    }
                $members += "`t`t`t</ul>`r`n"
                }
            $user = [PSCustomObject]@{
                GroupName = $g.Name
                Path = $path
                AccessType = $accessType
                Members = $members
            }
            $parsed += $user
            }
            return $parsed
        } catch {
            Write-Host $_
        }
    }
    end {

    }
}

#==========================================================================
# FUNCTION ConvertTo-ParsedUsers
#==========================================================================
Function ConvertTo-HtmlTable {
    <#
        .SYNOPSIS
        Send an e-mail to recipient
        .DESCRIPTION
        Send an e-mail to recipient
    #>
    [CmdletBinding()]
    Param(
        [Object]$header,
        [Object]$dataKeys,
        [Object]$dataValues
    )
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Send mail
        try {
            $table = "<table id=`"iTable`" cellpadding=`"2`" cellspacing=`"0`" border=`"1`">"

            <# Table Header #>
            $table += "`t<tr bgcolor=`"#000000`" style=`"color: #FFFFFF; height: 35px; border-collapse: collapse; border: 2px solid #000000;`">`n`r"
            ForEach($h in $header) {
                $table += "`t`t<td style=`"text-align: center; font-weight: 700;`">$($h)</td>`n`r"
            }
            $table += "`t</tr>`n`r"

            <# Table Body #>
            ForEach($d in $dataValues) {
                $table += "`t<tr>`n`r"
                    ForEach($k in $dataKeys){
                        $table += "`t`t<td class=`"td_$($k)`">$($d.$k)</td>`n`r"
                    }
                $table += "`t</tr>`n`r"
            }
            $table += "</table>`n`r"

            return $table
        } catch {
            Write-Host $_
        }
    }
    end {

    }
}
#==========================================================================


#==========================================================================
# FUNCTION Get-YaEmailList
#==========================================================================
Function Get-YaEmailList {
    <#
        .SYNOPSIS
        Get list of emails from Yandex via API
        .DESCRIPTION
        Get list of emails from Yandex via API
    #>
    begin {
        [string]$FunctionName = $PSCmdlet.MyInvocation.MyCommand.Name
    }

    process {
        # Request
        try {
            $config = Get-Content '.\config.json' | Out-String | ConvertFrom-Json

            $headers = @{Authorization = "Bearer " + $config.token}
            $uri = "https://api360.yandex.net/directory/v1/org/$($config.orgId)/users?perPage=200"
            $res = (Invoke-WebRequest -Headers $Headers -Uri $Uri).Content
            $users = (ConvertFrom-Json -InputObject $res).users
            $result = @()
            ForEach($user in $users) {
              $result += [pscustomobject]@{
                id = $user.id
                nickname = $user.nickname
                firstName = $user.name.first
                lastName = $user.name.last
                email = $user.email
                gender = $user.gender
                position = $user.position
                isEnabled = $user.isEnabled
                timezone = $user.timezone
                language = $user.language
                createdAt = $user.createdAt
                updatedAt = $user.updatedAt
              }
            }
            return $result
        } catch {
            Write-Host $_
        }
    }
    end {

    }
}