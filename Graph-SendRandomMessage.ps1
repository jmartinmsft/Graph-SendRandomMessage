<#//***********************************************************************
//
// Graph-SendRandomMessages.ps1
// Modified 2021/08/19
// Last Modifier:  Jim Martin
// Project Owner:  Jim Martin
// Version: v1.0
// Syntax for running this script:
//
// .\Graph-SendRandomMessages.ps1
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
[Parameter(Mandatory=$false)] [int] $NumberOfMessages=10,
[Parameter(Mandatory=$false)] [string]$UserPrincipalName
)
function CreateWord {
Param(
 [Parameter(Mandatory=$true)] [int]$LetterCount
)
	$Word = -join ((65..90) + (97..122) | Get-Random -Count $LetterCount | % {[char]$_})
	return $Word
}
function Write-Disclaimer{
Write-Host -ForegroundColor Yellow '//***********************************************************************'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// Copyright (c) 2018 Microsoft Corporation. All rights reserved.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR'
Write-Host -ForegroundColor Yellow '// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,'
Write-Host -ForegroundColor Yellow '// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE'
Write-Host -ForegroundColor Yellow '// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER'
Write-Host -ForegroundColor Yellow '// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,'
Write-Host -ForegroundColor Yellow '// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN'
Write-Host -ForegroundColor Yellow '// THE SOFTWARE.'
Write-Host -ForegroundColor Yellow '//'
Write-Host -ForegroundColor Yellow '//**********************************************************************​'
Start-Sleep -Seconds 2
}
$Modules = Get-Module
if ("ExchangeOnlineManagement" -notin  $Modules.Name) {
    try {Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -PSSessionOption $SessionOptions -ShowBanner:$False}
    catch {Write-Warning "Failed to connect to Exchange Online"; break }
}
Write-Host "Getting list of available mailboxes to send as..." -NoNewline -ForegroundColor Cyan
$Mailboxes = (Get-EXOMailbox -ResultSize unlimited | where {$_.Name -notlike "DiscoverySearch*"}).PrimarySmtpAddress
Write-Host "COMPLETE" -ForegroundColor Green
Write-Host "Getting a list of recpipients to send to..." -NoNewline -ForegroundColor Cyan
$Recipients = (Get-Recipient -ResultSize Unlimited | Where {($_.RecipientType -eq "UserMailbox" -and $_.Name -notlike "DiscoverySearch*") -or $_.RecipientType -eq "MailUser"}).PrimarySmtpAddress
Write-Host "COMPLETE" -ForegroundColor Green
#Change the AppId, AppSecret, and TenantId to match your registered application
$AppId = "15a65478-8dea-410b-b6c1-c1662692a63b"
$AppSecret = "_IcF_o8Agv2uXP_-r.I317g4bqHIg2V-E9"
$TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
#Build the URI for the token request
$Uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$Body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}
$TokenRequest = Invoke-WebRequest -Method Post -Uri $Uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
#Unpack the access token
$Token = ($TokenRequest.Content | ConvertFrom-Json).Access_Token
#Create headers to send Bearer auth header
$Headers = @{
    'Content-Type'  = "application\json"
    'Authorization' = "Bearer $Token" 
}
#HTML header with styles
$htmlhead="<html>
    <style>
    BODY{font-family: Arial; font-size: 10pt;}
	H1{font-size: 22px;}
	H2{font-size: 18px; padding-top: 10px;}
	H3{font-size: 16px; padding-top: 8px;}
    </style>"
# Mail.Send can send from any user. If you use an application access policy to restrict access to the API, make sure this user is included.
for ($i=1;$i -le $NumberOfMessages; $i++) {
    $pc = ($i/$NumberOfMessages)*100 
    Write-Progress -Activity "Spamming..." -CurrentOperation "$i of $NumberOfMessages complete" -Status "Please wait." -PercentComplete $pc
    $MsgFrom = Get-Random -Count 1 -InputObject $Mailboxes
    #$ccRecipient = "ronan@thejimmartin.com"
    #You can include an attachment if needed
    #$AttachmentFile = "C:\temp\WelcomeToOffice365ITPros.docx"
    #$ContentBase64 = [convert]::ToBase64String( [system.io.file]::readallbytes($AttachmentFile))
    $ToRecipients = Get-Random -Count (Get-Random -Minimum 1 -Maximum 10) -InputObject $Recipients
    [System.Collections.ArrayList]$RecipientList = New-Object System.Collections.ArrayList
    foreach($r in $ToRecipients){
        $RecipientList.Add(@{"emailAddress"=@{"address"=$r}}) | Out-Null
    }
    $MsgSubject = "Testing OAuth application to send email"
    #Create message body and properties and send
    #$htmlMsg = "just some basic text"
    [System.Collections.ArrayList]$Body = New-Object System.Collections.ArrayList
    $BodyWordCount = Get-Random -Minimum 5 -Maximum 500
    for($x=1; $x -le $BodyWordCount; $x++) {
        $BodyWord = CreateWord (Get-Random -Minimum 1 -Maximum 8)
        $Body.Add($BodyWord) | Out-Null
    }
    [string]$HtmlMsg = $Body
    $MessageParameters = @{
        "URI"         = "https://graph.microsoft.com/v1.0/users/$MsgFrom/sendMail"
        "Headers"     = $Headers
        "Method"      = "POST"
        "ContentType" = 'application/json'
        "Body" = (@{
            "message" = @{
            "subject" = $MsgSubject
            "body"    = @{
            "contentType" = 'HTML' 
            "content"     = $HtmlMsg }
            #"attachments" = @(
            #   @{
            #    "@odata.type" = "#microsoft.graph.fileAttachment"
            #    "name" = $AttachmentFile
            #    "contenttype" = "application/vnd.openxmlformats-officedocument.Wordprocessingml.document"
            #    "contentBytes" = $ContentBase64 } )  
            "toRecipients" = @($RecipientList)
            #@{
            #   "emailAddress" = @{"address" = $EmailRecipient }
            #} ) 
            #"ccRecipients" = @(
            # @{
            #   "emailAddress" = @{"address" = $ccRecipient1 }
            # } ,
            #  @{
            #   "emailAddress" = @{"address" = $ccRecipient2 }
            # } )       
            }
        }) | ConvertTo-Json -Depth 6
    }
    # Send the message
    Invoke-RestMethod @MessageParameters | Out-Null
}