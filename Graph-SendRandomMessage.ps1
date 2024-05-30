<#//***********************************************************************
//
// Graph-SendRandomMessages.ps1
// Modified 23 October 2023
// Last Modifier:  Jim Martin
// Project Owner:  Jim Martin
// Version: v20231023.1015
// Syntax for running this script:
//
// .\Graph-SendRandomMessages.ps1 -Sender jim@contoso.com -ToRecipients jeff@contoso.com
//
//***********************************************************************
//
// Copyright (c) 2018 Microsoft Corporation. All rights reserved.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
//**********************************************************************​
#>
param(
[Parameter(Mandatory=$false)] [int] $NumberOfMessages=3,
[Parameter(Mandatory=$false)] [string]$SenderAddress=$null,
[Parameter(Mandatory=$false)] $ToRecipients,
[Parameter(Mandatory=$false)] $CcRecipients,
[Parameter(Mandatory=$false)] $BccRecipients,
[Parameter(Mandatory=$false)] [switch]$IncludeAttachment,
[Parameter(Mandatory=$false)] [string]$AttachmentPath,
[Parameter(Mandatory=$false)] [switch]$RandomRecipients,
#>** OAUTH PARAMETERS START **#
[Parameter(Mandatory=$False,HelpMessage="The OAuthClientId specifies the client Id that this script will identify as.  Must be registered in Azure AD.")] [string]$OAuthClientId = "2e542266-3c04-4354-8965-aeafccd61976",
[Parameter(Mandatory=$False,HelpMessage="The OAuthTenantId specifies the tenant Id (application must be registered in the same tenant being accessed).")] [string]$OAuthTenantId = "9101fc97-5be5-4438-a1d7-83e051e52057",
[Parameter(Mandatory=$False,HelpMessage="The OAuthRedirectUri specifies the redirect Uri of the Azure registered application.")] [string]$OAuthRedirectUri = "msal9b38df47-ae02-4777-9edb-4ba2b727bcc4://auth",
[Parameter(Mandatory=$False,HelpMessage="The OAuthSecretKey specifies the secret key if using application permissions.")] [string]$OAuthSecretKey,
[Parameter(Mandatory=$False,HelpMessage="The OAuthCertificate specifies the certificate if using application permissions.  Certificate auth requires MSAL libraries to be available.")] $OAuthCertificate = $null
#>** OAUTH PARAMETERS END **#

)

function GetOAuthToken {
    if($null -notlike $OAuthCertificate) {
        $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -RedirectUri $RedirectUri -TenantId $OAuthTenantId -Scopes $Script:Scope -AzureCloudInstance AzurePublic -ClientCertificate (Get-Item Cert:\CurrentUser\My\$CertificateThumbprint)
    }
    else {
        $OAuthSecretKey = $OAuthSecretKey | ConvertTo-SecureString -Force -AsPlainText
        $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -ClientSecret $OAuthSecretKey -TenantId $OAuthTenantId -Scopes $Script:Scope -AzureCloudInstance AzurePublic
    }
    return $Script:OAuthToken.AccessToken
}

function GetRequestHeader{
    #Change the AppId, AppSecret, and TenantId to match your registered application
    #$AppId = "2e542266-3c04-4354-8965-aeafccd61976"
    #$TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
    
    #$Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    #$Body = @{
    #    client_id     = $AppId
    #    scope         = "https://graph.microsoft.com/.default"
    #    client_secret = $AppSecret
    #    grant_type    = "client_credentials"
    #}   
    #$Script:OAuthToken = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
    #$AccessToken = ($Script:OAuthToken.Content | ConvertFrom-Json).Access_Token
    #$script:OAuthTokenAcquireTime = [DateTime]::UtcNow
    #Create headers to send Bearer auth header
    $RequestHeaderInfo = @{
        "ExpiresIn" = $TokenExpiresIn
        "AuthHeaders" = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $AccessToken" 
        }
    }
    return $RequestHeaderInfo
}
#$Recipients = (Get-Mailbox -Filter "WindowsEmailAddress -notlike '*thejimmartin.onmicrosoft.com' -and RecipientTypeDetails -eq 'UserMailbox'").PrimarySmtpAddress
# Get OAuth token and create request header
#$Script:OAuthTokenAcquireTime = [DateTime]::UtcNow
$Script:Scope = "https://graph.microsoft.com/.default"

$Token = GetOAuthToken
$Headers = @{
    'Content-Type'  = "application\json"
    'Authorization' = "Bearer $Token"
}
$GetRandomSender = $true

#Get a list of words to use for subject and message body
if($null -eq $global:WordList){
    $global:WordList  = Get-Content .\words.txt
}

if($null -notlike $SenderAddress){
    $GetRandomSender = $false
}
Write-Host "Spamming to start now..." -ForegroundColor green
# Mail.Send can send from any user. If you use an application access policy to restrict access to the API, make sure this user is included.
for ($i=1;$i -le $NumberOfMessages; $i++) {
    $pc = ($i/$NumberOfMessages)*100 
    #$TokenExpiresInSeconds = ($Script:OAuthToken.ExpiresOn.UtcDateTime - [DateTime]::UtcNow).TotalSeconds
    Write-Progress -Activity "Spamming..." -CurrentOperation "$i of $NumberOfMessages from $SenderAddress complete" -Status "Please wait." -PercentComplete $pc
    if(($Script:OAuthToken.ExpiresOn.LocalDateTime).AddMinutes(-5) -le (Get-Date)) { 
        Write-Host "Renewing the token..." -ForegroundColor Yellow
        $Script:OAuthToken = Get-MsalToken -ClientId $OAuthClientId -TenantId $OAuthTenantId -Scopes $Script:Scope -ForceRefresh
        $Token = $Script:OAuthToken.AccessToken
        $Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token"
        } 
    }
    #Determine the sender
    if($GetRandomSender -eq $true) {
        if($null -eq $Global:Mailboxes) {
            $Global:Mailboxes = (Get-Mailbox -ResultSize unlimited | Where-Object {$_.PrimarySmtpAddress -notlike "*onmicrosoft*"}).PrimarySmtpAddress
        }
        $SenderAddress = Get-Random -InputObject $Global:Mailboxes -Count 1
    }
    #Write-Host $SenderAddress
    #Determine the list of recipients
    [System.Collections.ArrayList]$ToRecipientList = New-Object System.Collections.ArrayList
    if($RandomRecipients) {
        [int]$NumberOfRecipients = Get-Random -Minimum 1 -Maximum 5
        $ToRecipientsAddresses = Get-Random -InputObject $ToRecipients -Count $NumberOfRecipients
    }
    else {
        $ToRecipientsAddresses = $ToRecipients
    }
    foreach($r in $ToRecipientsAddresses){
        $ToRecipientList.Add(@{"emailAddress"=@{"address"=$r}}) | Out-Null
    }
    if($RandomRecipients -and $null -ne $CcRecipients) {
        [int]$NumberOfRecipients = Get-Random -Minimum 1 -Maximum 5
        $CcRecipientsAddresses = Get-Random -InputObject $CcRecipients -Count $NumberOfRecipients
    }
    else {
        $CcRecipientsAddresses = $CcRecipients
    }
    if($null -ne $CcRecipientsAddresses) {
        [System.Collections.ArrayList]$CcRecipientList = New-Object System.Collections.ArrayList
        foreach($r in $CcRecipientsAddresses){
            $CcRecipientList.Add(@{"emailAddress"=@{"address"=$r}}) | Out-Null
        }
    }
    if($RandomRecipients -and $null -ne $BccRecipients) {
        [int]$NumberOfRecipients = Get-Random -Minimum 1 -Maximum 5
        $BccRecipientsAddresses = Get-Random -InputObject $BccRecipients -Count $NumberOfRecipients
    }
    else {
        $BccRecipientsAddresses = $BccRecipients
    }
    if($null -ne $BccRecipientsAddresses) {
        [System.Collections.ArrayList]$BccRecipientList = New-Object System.Collections.ArrayList    
        foreach($r in $BccRecipientsAddresses){
            $BccRecipientList.Add(@{"emailAddress"=@{"address"=$r}}) | Out-Null
        }
    }
    # Generate the message subject
    [string]$MsgSubject = $null
    [int]$NumberOfWords = Get-Random -Minimum 1 -Maximum 6
    [string]$subject = Get-Random -InputObject $global:WordList -Count $NumberOfWords
    [string]$MsgSubject = "GraphTest " + $subject
    # Generate the message body
    [int]$NumberOfWords = Get-Random -Minimum 1 -Maximum 500
    [string]$Body = Get-Random -InputObject $global:WordList -Count $NumberOfWords
    [string]$HtmlMsg = $Body
    # Generate the Graph API request message body
    $MessageBody = (@{
        "message" = @{
            "subject" = $MsgSubject
            "body"    = @{
                "contentType" = 'HTML' 
                "content"     = $HtmlMsg
            }
            "toRecipients" = @($ToRecipientList)
        }
    })
    if($CcRecipients){
        $MessageBody.message.add("ccRecipients", @($CcRecipientList))
    }
    if($BccRecipients){
        $MessageBody.message.add("bccRecipients", @($BccRecipientList))
    }
    # Prepare the Graph API request

    $MessageParameters = @{
        "URI"         = "https://graph.microsoft.com/v1.0/users/$($SenderAddress)/sendMail"
        "Headers"     = $Headers
        "Method"      = "POST"
        "ContentType" = 'application/json'
        "Body" = $MessageBody | ConvertTo-Json -Depth 6
    }
    # Send the message
    try {
        $Error.Clear()
        Invoke-RestMethod @MessageParameters | Out-Null
    }
    catch {
        if($Error[0] -like "*underlying connection was closed*") {
            Invoke-RestMethod @MessageParameters | Out-Null
        }
    }
    $WaitTime = (Get-Random -Minimum 1 -Maximum 5)
    $NextSend = (Get-Date).AddSeconds($WaitTime)
    #Write-Host "Sending the next message at $NextSend"
    Start-Sleep -Seconds $WaitTime
}