[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $Messenger
)

#Requires -module PSTeams
import-module PSTeams

[string]$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
[string]$scriptConfigPath = "$scriptPath\config.json" 
$scriptConfig = Get-Content -raw -Path  $scriptConfigPath | ConvertFrom-Json

#Tenant configuration
[string]$clientId = $scriptConfig.ScriptMainConfig.TenantConfig.clientId
[string]$tenantId = $scriptConfig.ScriptMainConfig.TenantConfig.tenantId
[string]$clientSecret = $scriptConfig.ScriptMainConfig.TenantConfig.clientSecret
[string]$graphUrl = $scriptConfig.ScriptMainConfig.TenantConfig.graphUrl

#Time frame configuration (in minutes)
[int]$RefreshTime = $scriptConfig.ScriptMainConfig.RefreshTime

#Messengers config
[string]$TeamsID = $scriptConfig.ScriptMainConfig.TeamsConfig.teamsWebhookID
[string]$tokenTelegram = $scriptConfig.ScriptMainConfig.TelegramConfig.telegramBotToken
[string]$telegramChatID = $scriptConfig.ScriptMainConfig.TelegramConfig.telegramChannelId

Function Get-GraphToken(){
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $AppId,
        $AppSecret,
        $TenantID
    )

    $AuthUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $Scope = "https://manage.office.com/.default"

    $Body = @{
        client_id = $AppId
            client_secret = $AppSecret
            scope = $Scope
            grant_type = 'client_credentials'
    }

    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        Body = $Body
        Uri = $AuthUrl
    }

    try {
        $Token = Invoke-RestMethod @PostSplat -ErrorAction Stop
        Write-Verbose "Token successfully generated"
        return $Token
    }
    catch {
        Write-Warning "Exception was caught: $($_.Exception.Message)" 
    }
}

Function Get-GraphResult () {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $Url, 
        $Token, 
        $Method
    )

    $Header = @{
        Authorization = "$($Token.token_type) $($Token.access_token)"
    }

    $PostSplat = @{
        ContentType = 'application/json'
        Method = $Method
        Header = $Header
        Uri = $Url
    }

    try {
        Invoke-RestMethod @PostSplat -ErrorAction Stop
    }
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        break
    }
}

Function Get-M365Health(){
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $graphUrl,
        $tenantId,
        $token
    )

    $graphApiVersion = "v1.0"
    $MC_resource = "ServiceComms/Messages" 
    $uri = "$graphUrl/$graphApiVersion/$($tenantId)/$MC_resource"
    write-host "$uri"
    $Method = "GET"

    try {
            $Result = Get-GraphResult -Url $uri -Token $Token -Method $Method
            Write-Verbose "Messages successfully collected"
            return $Result.value
    }
    catch {
            $ex = $_.Exception
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd();
            Write-Errpr "Response content:`n$responseBody" -f Red
            Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
            break
            Write-Error "Can't get new messaegs"
    }
}

function Send-FormattedTelegramMessage {

    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true)]
        $message,
        [string]$tokenTelegram,
        [string]$chatID
    )
    #Formatting message 
    
    $MessageID = $message.Id
    $WorkloadDisplayName = $message.WorkloadDisplayName
    $Status = $message.Status
    #$MessageTitle = $message.Title

    $MessageDetails = $message.Messages[$message.Messages.Count-1].MessageText

    $MessageDetailsForTgm = $MessageDetails

    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Title\:','*Title:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'More info\:','*More info:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Current status\:','*Current status:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Root cause\:','*Root cause:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Next update by\:','*Next update by:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Scope of impact\:','*Scope of impact:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Next steps\:','*Next steps:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Start time\:','*Start time:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'End time\:','*End time:**'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'User impact\:','*User impact:*'
    $MessageDetailsForTgm = $MessageDetailsForTgm -replace 'Final status\:','*Final status:*'
    #$MessageDetailsForTgm = $MessageDetailsForTgm -split "`n" | Select-object -Skip 1 #needs to remove duplicaded Title

    $NotificationTitle = "$WorkloadDisplayName : $Status ($MessageID)"

    $FinalMessageForTgm = "*$NotificationTitle* `n$MessageDetailsForTgm"

    #Sending message
    $URL_set = "https://api.telegram.org/bot$tokenTelegram/sendMessage"
      

    $body = @{
        text = $FinalMessageForTgm
        parse_mode = "markdown"
        chat_id = $chatID
    }

    $messageJson = $body | ConvertTo-Json

    try {
        Invoke-RestMethod $URL_set -Method Post -ContentType 'application/json; charset=utf-8' -Body $messageJson
        Write-Verbose "Message has been sent"
    }
    catch {
        Write-Error "Can't sent message"
    }
    
}

function Send-FormattedTeamsMessage {

    [CmdletBinding()]

    param (
        [Parameter(Mandatory=$true)]
        $message,
        [string]$weebhookid
    )
    #Format Title
    $MessageID = $message.Id
    $WorkloadDisplayName = $message.WorkloadDisplayName
    $Status = $message.Status
    $MessageTitle = $message.Title
    $MessageDetails = $message.Messages[$message.Messages.Count-1].MessageText
    
    #formatting block
    $MessageDetails = $MessageDetails -replace 'Title\:','**Title:**'
    $MessageDetails = $MessageDetails -replace 'More info\:','**More info:**'
    $MessageDetails = $MessageDetails -replace 'Current status\:','**Current status:**'
    $MessageDetails = $MessageDetails -replace 'Root cause\:','**Root cause:**'
    $MessageDetails = $MessageDetails -replace 'Next update by\:','**Next update by:**'
    $MessageDetails = $MessageDetails -replace 'Scope of impact\:','**Scope of impact:**'
    $MessageDetails = $MessageDetails -replace 'Next steps\:','**Next steps:**'
    $MessageDetails = $MessageDetails -replace 'Start time\:','**Start time:**'
    $MessageDetails = $MessageDetails -replace 'End time\:','**End time:**'
    $MessageDetails = $MessageDetails -replace 'User impact\:','**User impact:**'
    $MessageDetails = $MessageDetails -replace 'Final status\:','**Final status:**'

    $NotificationTitle = "$WorkloadDisplayName : $MessageTitle - $Status ($MessageID)"

    if ($status -eq 'Service Restored') {
        Send-TeamsMessage -Uri $weebhookid -MessageTitle $NotificationTitle -MessageText $MessageDetails -Color Green                
    }elseif ($status -eq 'Service degradation' -or $status -eq 'Service Interruption') {
        Send-TeamsMessage -Uri $weebhookid -MessageTitle $NotificationTitle -MessageText $MessageDetails -Color Red
    }else {
        Send-TeamsMessage -Uri $weebhookid -MessageTitle $NotificationTitle -MessageText $MessageDetails -Color YellowGreen
    }

}

Write-Verbose "Getting the M365 Graph Token"

try {
    $GraphToken = Get-GraphToken -AppId $clientId -AppSecret $clientSecret -TenantID $tenantId -ErrorAction Stop
    Write-Verbose "Token successfully issued"
}
catch {
    Write-Error "Can't get the token!"
    break
}

Write-Verbose "Collecting all statuses"

try {
    $ServiceHealth = Get-M365Health -graphUrl $graphUrl -tenantId $tenantId -token $GraphToken -ErrorAction Stop
}
catch {
    Write-Error "Can't collect messages"
    break    
}

#manipulations with $time for schedulling and checking the updates
$TimeRange = -$($RefreshTime)
$TimeCheckPoint = Get-Date
Write-Verbose "Time Chekpoint: $TimeCheckPoint"
$LastRunTime = $($TimeCheckPoint.AddMinutes($TimeRange)).ToUniversalTime()
Write-Verbose "Last run time: $LastRunTime"

$NewMessages = $ServiceHealth | Where-Object {$_.MessageType -eq "Incident" -and $($(get-date $_.LastUpdatedTime).ToUniversalTime()) -gt $(get-date $LastRunTime)}
$NewMessagesCount = $NewMessages.count

if ($NewMessages) {
        Write-Verbose "There are $NewMessagesCount new messages"
        foreach($NewMessage in $NewMessages){
            switch ($Messenger) {
                Teams {
                    Send-FormattedTeamsMessage -message $NewMessage -weebhookid $TeamsID
                }
                Telegram{
                    Send-FormattedTelegramMessage -message $NewMessage -tokenTelegram $tokenTelegram -chatID $telegramChatID
                }
                All {
                    Send-FormattedTeamsMessage -message $NewMessage -weebhookid $TeamsID
                    Send-FormattedTelegramMessage -message $NewMessage -tokenTelegram $tokenTelegram -chatID $telegramChatID
                }
                Default {
                    Write-Warning "Messenger was not choosen"
                }
            }
        }
}else {
    Write-Verbose "There are no new messages to send"
}



