<#
  .SYNOPSIS
  Basic POC/Example - Threat Submission using Security Graph API with TABL addition
  
  Permission needed for submission --> Application -ThreatSubmission.ReadWrite.All
  https://learn.microsoft.com/en-us/graph/api/security-emailthreatsubmission-post-emailthreats?view=graph-rest-beta&tabs=http
  
  Permission needed to get the message id as refrence for submission --> Application - Mail.ReadBasic.All OR Mail.Read
  https://learn.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http

  Additinal modules needed
  * MSAL.PS Powershell module as prerequisite. https://www.powershellgallery.com/packages/MSAL.PS/4.37.0.0
  * JWT https://www.powershellgallery.com/packages/JWTDetails/1.0.2

  .DESCRIPTION
  Microsoft Defender for Office 365 Admin Submission using Graph API with application Permission
  It´s a simple proof of concept with no further error managment.

#>

$tenantID = ''
$clientID = ''
$clientSecret = ConvertTo-SecureString '' -AsPlainText -Force

# Using MSAL.PS powershell library to get the access token
Write-Host "[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Getting accesstoken using clientsecret." -ForegroundColor DarkGreen
$authResult = Get-MsalToken -ClientId $clientID -TenantId $tenantID -ClientSecret $clientSecret -ForceRefresh
$accessToken = $authResult.AccessToken
# Using jwtdetails module to get the details of the access token
Write-Host "[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Accesstoken details" -ForegroundColor Yellow
$accessToken | Get-JWTDetails | select-object aud,app_displayname,roles
$accessTokenSecureString = ConvertTo-SecureString $accessToken -AsPlainText -Force

$headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}

$Mailbox = ""
$MessageID = "''"
$MessageIDQueryURL = "https://graph.microsoft.com/v1.0/users/{0}/messages?`$filter=(internetMessageId eq {1})" -f $Mailbox,$MessageID
Write-Host ("[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Getting message id for internetMessageId {0} in mailbox {1}." -f $MessageID,$Mailbox) -ForegroundColor DarkGreen
$id = (Invoke-RestMethod -Headers $headers -Uri $MessageIDQueryURL -Method Get).value.id
Write-Host ("[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Message id resolved to {0}." -f $id) -ForegroundColor DarkGreen


$GraphUrl="https://graph.microsoft.com/beta/security/threatSubmission/emailThreats"
$messageURL = "https://graph.microsoft.com/v1.0/users/{0}/messages/{1}" -f $Mailbox,$id

$allowAllow = $true
if ( $allowAllow ) {
  $expirationDate = (get-date).adddays(+30) | get-date -Format o
  $submissionCategory = "notjunk" 
  $bodyJSON = [PSCustomObject]@{
    '@odata.type' = '#microsoft.graph.security.emailUrlThreatSubmission'
    category = $submissionCategory
    recipientEmailAddress = $Mailbox
    messageUrl = $messageURL
    tenantAllowOrBlockListAction =
      @{
        action = 'allow'
        expirationDateTime = $expirationDate
        note = 'temporal allow the url/attachment/sender in the email - API done'
      }
  } | ConvertTo-Json

} else {
  $submissionCategory = "phishing" 
  $bodyJSON = [PSCustomObject]@{
    '@odata.type' = '#microsoft.graph.security.emailUrlThreatSubmission'
    category = $submissionCategory
    recipientEmailAddress = $Mailbox
    messageUrl = $messageURL   
  } | ConvertTo-Json
}

Write-Host ("[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Submitting email threat with category {0} for mailbox {1}." -f $submissionCategory,$Mailbox) -ForegroundColor DarkGreen
$err=$null
try{ $Submissionresult = Invoke-WebRequest -Uri $GraphURL -Headers $headers -Body $bodyJSON -Method POST -ContentType 'application/json' -ErrorVariable RespErr } catch {$err=$_.Exception}
if ( $err ) {
  Write-Host "[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Ran into an exception while submitting the email threat." -ForegroundColor Red
} else {
  Write-Host ("[$((Get-Date).ToString("yyyy/MM/dd HH:mm:ss.fff"))] Submission completed with status code {0}." -f $Submissionresult.StatusCode) -ForegroundColor DarkGreen
}
$err | Get-Member -MemberType Property

