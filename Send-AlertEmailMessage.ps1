 
 function Send-AlertMessage {
 [CmdletBinding()]
    param(
        [string]$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        [string]$ClientSecret = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        [string]$TenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        [string]$TemplatePath = "$($PSScriptRoot)\Template",
        [string]$Template = "Template.html",
        [string]$TemplateColor = "#CC0000",
        [string]$ComputerName = "$($ENV:COMPUTERNAME)",
        [string]$EventName = "Data Acquisition Task",
        [string]$TaskName = "Azure AD Data Aquisition Log",
        [string]$TaskStatus = "Failed",
        [string]$MailSender = "SmtpEmail@domain.com",
        [Parameter(Mandatory=$True,
            HelpMessage="Recipient Email Address (e.g username@domain.com)")]
        [array]$EmailAddress,
        [Parameter(HelpMessage="Carbon Copy Email Address (e.g username@domain.com)")]
        [array]$CCAddress,
        [Parameter(HelpMessage="Blind Carbon Copy Email Address (e.g username@domain.com)")]
        [array]$BCCAddress,
        [string]$Subject,
        [string]$Body
    )

    if ($EmailAddress) {
        $EmailinJSON = $EmailAddress | %{'{"EmailAddress": {"Address": "'+$_+'"}},'}
        $EmailinJSON = ([string]$EmailinJSON).Substring(0, ([string]$EmailinJSON).Length - 1)
        $EmailinJSON = "`"toRecipients`": [ $EmailinJSON ]"
    }
    if($CCAddress) { 
        $CCinJSON = $CCAddress | %{'{"EmailAddress": {"Address": "'+$_+'"}},'}
        $CCinJSON = ([string]$CCinJSON).Substring(0, ([string]$CCinJSON).Length - 1)
        $CCinJSON = "`"ccRecipients`": [ $CCinJSON ]"
        }
    if($BCCAddress) { 
        $BCCinJSON = $BCCAddress | %{'{"EmailAddress": {"Address": "'+$_+'"}},'}
        $BCCinJSON = ([string]$BCCinJSON).Substring(0, ([string]$BCCinJSON).Length - 1)
        $BCCinJSON = "`"bccRecipients`": [ $BCCinJSON ]"
    }

    If (($EmailAddress -and $CCAddress -and !$BCCAddress) -or ($EmailAddress -and !$CCAddress -and $BCCAddress)) {
        $EmailinJSON = "$($EmailinJSON),"
    } ElseIf (!$EmailAddress -and $CCAddress -and $BCCAddress) {
        $CCinJSON = "$($CCinJSON),"
    } ElseIf ($EmailAddress -and $CCAddress -and $BCCAddress) {
        $EmailinJSON = "$($EmailinJSON),"
        $CCinJSON = "$($CCinJSON),"
    }

    $EmailRecipients = "$($EmailinJSON)
                        $($CCinJSON)
                        $($BCCinJSON)"

    $EventStatus = $TaskStatus
    $EventColor = $TemplateColor
    $StatusColor = $TemplateColor
    $Date = (Get-Date -UFormat "%m-%d-%Y")
    $EventDate = $Date
    $Time = Get-Date -Format "hh:mm tt"
    $TimeString  = $Time.Substring(0,1)
    If ($TimeString -eq 0) {
        $EventTime = $Time.Substring(1,7)
    } Else { 
        $EventTime = $Time
    }

    If ($TemplatePath -and $Template) {
        Push-Location $TemplatePath

        iex ('$Body = @"' + "`n" + ([string]::join("`n",(Get-Content $Template))) + "`n" + '"@')
        $Body = $Body.Replace('"', "'")

        Pop-Location
    }
    #Connect to GRAPH API
    $TokenBody = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $ClientId
        Client_Secret = $ClientSecret
    }
    $TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $TokenBody
    $Headers = @{
        "Authorization" = "Bearer $($TokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    #Send Mail    
$URLSend = "https://graph.microsoft.com/v1.0/users/$MailSender/sendMail"
$BodyJsonSend = @"
                    {
                        "message": {
                            "subject": "$($Subject)",
                            "Importance": "High",
                            "body": {
                                "contentType": "HTML",
                                "content": "$($Body)"
                            },
                            $EmailRecipients
                        },
                        "saveToSentItems": "false"
                   }
"@
$BodyJsonSend = $BodyJsonSend -replace ("(?m)^\s*`r`n",'').trim()

Write-Host `n"$($BodyJsonSend)"`n

Invoke-RestMethod -Method POST -Uri $URLSend -Headers $Headers -Body $BodyJsonSend
Write-Host `n

}

$Recipients = @()
$CCRecipients = @()
$BCCRecipients = @()

$Recipients += "Recipient@domain.com"
#$CCRecipients = ("CCRecipient1@domain.com","CCRecipient2@domain.com")
#$BCCRecipients = ("BCCRecipient1@domain.com","BCCRecipient1@domain.com")

Send-AlertMessage -EmailAddress $Recipients -CCAddress $CCRecipients -BCCAddress $BCCRecipients