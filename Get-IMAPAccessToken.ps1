
<#
  .SYNOPSIS
  Allows IMAP OAuth testing with Office 365.
  Please install MSAL.PS Powershell module as prerequisite. 
  https://github.com/DanijelkMSFT/ThisandThat/blob/main/Get-IMAPAccessToken.ps1
  Refercing article with more insides 
  https://www.linkedin.com/pulse/start-using-oauth-office-365-popimap-authentication-danijel-klaric
  https://techcommunity.microsoft.com/t5/exchange-team-blog/announcing-oauth-2-0-client-credentials-flow-support-for-pop-and/ba-p/3562963


  .DESCRIPTION
  The function helps admins to test their IMAP OAuth Azure Application, 
  with Interactive user login und providing or the lately released client credential flow
  using the right formatting for the XOAuth2 login string.
  After successful logon, a simple IMAP folder listing is done, in addition it also allows to 
  test shared mailbox acccess for users if fullaccess has been provided. 
  
  Using Windows Powershell allows MSAL to cache the access+refresh token on disk for further executions for interactive login scenario.
  It´s a simple proof of concept with no further error managment.

  .PARAMETER tenantID
  Specifies the target tenant.

  .PARAMETER clientId
  Specifies the ClientID/ApplicationID of the registered Azure AD Application with needed IMAP Graph permissions

  .PARAMETER clientsecret
  Specifies the ClientSecret of the registered Azure AD Application for client credential flow

  .PARAMETER targeMailbox
  Specifies the primary emailaddress of the targetmailbox which should be accessed by service principal which has fullaccess to for client credential flow

  .PARAMETER redirectUri
  Specifies the redirectUri of the registered Azure AD Application for authorization code flow (interactive flow)

  .PARAMETER LoginHint
  Specifies the Userprincipalname of the logging in user for authorization code flow (interactive flow)

  .PARAMETER SharedMailbox (optinal)
  Specifies the primary emailaddress of the Sharedmailbox logged in user has fullaccess to for authorization code flow (interactive flow)

  .EXAMPLE
  PS> .\Get-IMAPAccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com"

  .EXAMPLE
  PS> .\Get-IMAPAccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com" -SharedMailbox "SharedMailbox@contoso.com"

  .EXAMPLE
  PS> .\Get-IMAPAccessToken.ps1 -tenantID "" -clientId "" -redirectUri "https://localhost" -LoginHint "user@contoso.com" -Verbose

  .EXAMPLE
  .\Get-IMAPAccessToken.ps1 -tenantID "" -clientId "" -clientsecret '' -targetMailbox "TargetMailbox@contoso.com"

  .EXAMPLE
  .\Get-IMAPAccessToken.ps1 -tenantID "" -clientId "" -clientsecret '' -targetMailbox "TargetMailbox@contoso.com" -Verbose

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)][string]$tenantID,
    [Parameter(Mandatory = $true)][String]$clientId,
    
    [Parameter(Mandatory = $true,ParameterSetName="authorizationcode")][String]$redirectUri,
    [Parameter(Mandatory = $true,ParameterSetName="authorizationcode")][String]$LoginHint,
    [Parameter(Mandatory = $false,ParameterSetName="authorizationcode")][String]$SharedMailbox,

    [Parameter(Mandatory = $true,ParameterSetName="clientcredentials")][String]$clientsecret,    
    [Parameter(Mandatory = $true,ParameterSetName="clientcredentials")][String]$targetMailbox
)

function Test-IMAPXOAuth2Connectivity {
# get Accesstoken via user authentication and store Access+Refreshtoken for next attempts
if ( $redirectUri ){
    $MsftPowerShellClient = New-MsalClientApplication -ClientId $clientID -TenantId $tenantID -RedirectUri $redirectURI  | Enable-MsalTokenCacheOnDisk -PassThru
    try {
        $authResult = $MsftPowerShellClient | Get-MsalToken -LoginHint $LoginHint -Scopes 'https://outlook.office365.com/.default'
    }
    catch  {
        Write-Host "Ran into an exception while getting accesstoken" -ForegroundColor Red
        $_.Exception.Message
        $_.FullyQualifiedErrorId
        break
    }
}

if ( $clientsecret ){
    $SecuredclientSecret = ConvertTo-SecureString $clientsecret -AsPlainText -Force
    $MsftPowerShellClient = New-MsalClientApplication -ClientId $clientID -TenantId $tenantID -ClientSecret $SecuredclientSecret 
    try {
        $authResult = $MsftPowerShellClient | Get-MsalToken -Scopes 'https://outlook.office365.com/.default'
    }
    catch  {
        Write-Host "Ran into an exception while getting accesstoken" -ForegroundColor Red
        $_.Exception.Message
        $_.FullyQualifiedErrorId
        break
    }
}


$accessToken = $authResult.AccessToken
$username = $authResult.Account.Username

# build authentication string with accesstoken and username like documented here
# https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#authenticate-connection-requests

# in the case if client credential usage we need to add the target mailbox like shared mailbox access
if ( $targetMailbox) { $SharedMailbox = $targetMailbox }

if ( $SharedMailbox ) {
    $b="user=" + $SharedMailbox + "$([char]0x01)auth=Bearer " + $accessToken + "$([char]0x01)$([char]0x01)"
    Write-Host "Accessing Sharedmailbox - $SharedMailbox - with Accesstoken of User $userName." -ForegroundColor DarkGreen
} else {
        $b="user=" + $userName + "$([char]0x01)auth=Bearer " + $accessToken + "$([char]0x01)$([char]0x01)"
        }

$Bytes = [System.Text.Encoding]::ASCII.GetBytes($b)
$POPIMAPLogin =[Convert]::ToBase64String($Bytes)

Write-Verbose "SASL XOAUTH2 login string $POPIMAPLogin"

# connecting to Office 365 IMAP Service
Write-Host "Connect to Office 365 IMAP Service." -ForegroundColor DarkGreen
$ComputerName = 'Outlook.office365.com'
$Port = '993'
    try {
        $TCPConnection = New-Object System.Net.Sockets.Tcpclient($($ComputerName), $Port)
        $TCPStream = $TCPConnection.GetStream()
        try {
            $SSLStream  = New-Object System.Net.Security.SslStream($TCPStream)
            $SSLStream.ReadTimeout = 5000
            $SSLStream.WriteTimeout = 5000     
            $CheckCertRevocationStatus = $true
            $SSLStream.AuthenticateAsClient($ComputerName,$null,[System.Security.Authentication.SslProtocols]::Tls12,$CheckCertRevocationStatus)
        }
        catch  {
            Write-Host "Ran into an exception while negotating SSL connection. Exiting." -ForegroundColor Red
            $_.Exception.Message
            break
        }
    }
    catch  {
    Write-Host "Ran into an exception while opening TCP connection. Exiting." -ForegroundColor Red
    $_.Exception.Message
    break
    }    

    # continue if connection was successfully established
    $SSLstreamReader = new-object System.IO.StreamReader($sslStream)
    $SSLstreamWriter = new-object System.IO.StreamWriter($sslStream)
    $SSLstreamWriter.AutoFlush = $true
    $SSLstreamReader.ReadLine()

    Write-Host "Authenticate using XOAuth2." -ForegroundColor DarkGreen
    # authenticate and check for results
    $command = "A01 AUTHENTICATE XOAUTH2 {0}" -f $POPIMAPLogin
    Write-Verbose "Executing command -- $command"
    $SSLstreamWriter.WriteLine($command) 
    #respose might take longer sometimes
    while (!$ResponseStr ) { 
        try { $ResponseStr = $SSLstreamReader.ReadLine() } catch { }
    }

    if ( $ResponseStr -like "*OK AUTHENTICATE completed.") 
    {
        $ResponseStr
        Write-Host "Getting mailbox folder list as authentication was successfull." -ForegroundColor DarkGreen
        $command = 'A01 LIST "" *'
        Write-Verbose "Executing command -- $command"
        $SSLstreamWriter.WriteLine($command) 

        $done = $false
        $str = $null
        while (!$done ) {
            $str = $SSLstreamReader.ReadLine()
            if ($str -like "* OK LIST completed.") { $str ; $done = $true } 
            elseif ($str -like "* BAD User is authenticated but not connected.") { $str; "Causing Error: IMAP protcol access to mailbox is disabled or permission not granted for client credential flow. Please enable IMAP protcol access or grant fullaccess to service principal."; $done = $true} 
            else { $str }
        }

        Write-Host "Logout and cleanup sessions." -ForegroundColor DarkGreen
        $command = 'A01 Logout'
        Write-Verbose "Executing command -- $command"
        $SSLstreamWriter.WriteLine($command) 
        $SSLstreamReader.ReadLine()

    } else {
        Write-host "ERROR during authentication $ResponseStr" -Foregroundcolor Red
    }

    # Session cleanup
    if ($SSLStream) {
        $SSLStream.Dispose()
    }
    if ($TCPStream) {
        $TCPStream.Dispose()
    }
    if ($TCPConnection) {
        $TCPConnection.Dispose()
    }
}

#check for needed msal.ps module
if ( !(Get-Module msal.ps -ListAvailable) ) { Write-Host "MSAL.PS module not installed, please check it out here https://www.powershellgallery.com/packages/MSAL.PS/" -ForegroundColor Red; break}

# execute function
Test-IMAPXOAuth2Connectivity

