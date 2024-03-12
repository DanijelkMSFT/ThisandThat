<#
  .SYNOPSIS
  Basic POC/Example - Threat Submission using Security Graph API with TABL addition
  https://learn.microsoft.com/en-us/graph/api/security-emailthreatsubmission-post-emailthreats?view=graph-rest-beta&tabs=http
  Permission needed --> Application	ThreatSubmission.ReadWrite.All

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
$authResult = Get-MsalToken -ClientId $clientID -TenantId $tenantID -ClientSecret $clientSecret -ForceRefresh
$accessToken = $authResult.AccessToken
# Using jwtdetails module to get the details of the access token
$accessToken | Get-JWTDetails | select-object aud,app_displayname,roles
$accessTokenSecureString = ConvertTo-SecureString $accessToken -AsPlainText -Force

$headers= @{"Content-Type" = "application/json" ; "Authorization" = "Bearer " + $accessToken}

$Mailbox = ""
$MessageID = "''"
$MessageIDQueryURL = "https://graph.microsoft.com/v1.0/users/{0}/messages?`$filter=(internetMessageId eq {1})" -f $Mailbox,$MessageID
$id = (Invoke-RestMethod -Headers $headers -Uri $MessageIDQueryURL -Method Get).value.id


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

try{ $Submissionresult = Invoke-WebRequest -Uri $GraphURL -Headers $headers -Body $bodyJSON -Method POST -ContentType 'application/json' -ErrorVariable RespErr } catch {$err=$_.Exception}
$err | Get-Member -MemberType Property

