<#
  .SYNOPSIS
  Allows Advanced Hunting queries using Microsoft 365 Defender Advanced hunting API.
  Please install 
  * MSAL.PS Powershell module https://www.powershellgallery.com/packages/MSAL.PS/
  * JWTDetails Powershell module https://www.powershellgallery.com/packages/JWTDetails/
  as prerequisite. 
  
  https://github.com/DanijelkMSFT/ThisandThat/blob/main/Query-M365DAdvancedHuntingAPI.ps1

  .DESCRIPTION
  It´s a simple proof of concept with no further error managment, for lightweight testing of AH API access before implementing API access using Azure Logicapps like described here
  add LinkedIn article URL HERE

  .PARAMETER tenantID
  Specifies the target tenant.

  .PARAMETER clientId
  Specifies the ClientID/ApplicationID of the registered Azure AD Application with needed AH permissions

  .PARAMETER clientSecret
  Specifies the ClientSercet of the registered Azure AD Application 

  .EXAMPLE
   PS>.\Query-M365DAdvacnedHuntingAPI.ps1 -tenantID "" -clientId "" -clientSecret ''
   
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$tenantID,
    [Parameter(Mandatory = $true)][String]$clientId,
    [Parameter(Mandatory = $true)][String]$clientSecret
)

function Query-M365DAdvancedHuntingAPI {

  #Import-Module MSAL.PS
  $authResult = Get-MsalToken -ClientId $clientID -TenantId $tenantID -ClientSecret $clientSecretSecureString -Scopes 'https://api.security.microsoft.com/.default'
  $accessToken = $authResult.AccessToken

  #Write-Host "Accesstoken details" -ForegroundColor Yellow
  $accessToken | Get-JWTDetails | Select-Object aud,app_displayname,roles,timeToExpiry

  $headers = @{ 
      "Authorization" = ("Bearer {0}" -f $accesstoken);
      "Content-Type" = "application/json";
  }

  $URL = "https://api.security.microsoft.com/api/advancedhunting/run"

  $Query = @{
      Query="EmailEvents 
      | where Timestamp > ago(30d)
      | where EmailDirection == `"Intra-org`"
      | where (ThreatTypes == `"Phish`") and (DeliveryAction == `"Delivered`" and DeliveryLocation == `"Inbox/folder`")
      | join kind=inner UrlClickEvents on `$left.NetworkMessageId==`$right.NetworkMessageId 
      | project-rename ClickTimeStamp = Timestamp1 
      | project-rename EmailReceiveTimeStamp = Timestamp 
      | project EmailReceiveTimeStamp,ClickTimeStamp,AccountUpn,RecipientEmailAddress,ActionType,ConfidenceLevel,Url,UrlChain,DetectionMethods
      | limit 1
      "
  }
  $body = $Query | ConvertTo-Json
  $response=$null
  $response = Invoke-RestMethod -Uri $URL -Method Post -Body $body -ContentType 'application/json' -Headers $headers
  $response.Results[0] | Out-String
}

$clientSecretSecureString = ConvertTo-SecureString $clientSecret -AsPlainText -Force
Query-M365DAdvancedHuntingAPI -tenantID $tenantID -clientId $clientId -ClientSecret $clientSecretSecureString

