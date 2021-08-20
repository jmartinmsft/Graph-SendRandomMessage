<#//***********************************************************************
//
// Graph-SendMail.ps1
// Modified 2021/08/18
// Last Modifier:  Jim Martin
// Project Owner:  Jim Martin
// Version: v1.0
//Syntax for running this script:
//
// .\Graph-SendMail.ps1
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
function Write-Disclaimer {
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
# Get Graph access token - change these values for the app you use.
$AppId = "4b7233c6-c05d-4d08-9d37-412bd260bed6"
$AppSecret = "P54cAXb.4Vf.v.bzp3jJQIu_7g8Y6NLX-t"
$TenantId = "9101fc97-5be5-4438-a1d7-83e051e52057"
# Construct URI and body needed for authentication
$Uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$Body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $AppSecret
    grant_type    = "client_credentials"
}
$TokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
# Unpack Access Token
$Token = ($TokenRequest.Content | ConvertFrom-Json).access_token
$Headers = @{
            'Content-Type'  = "application\json"
            'Authorization' = "Bearer $Token" }
# Mail.Send can send from any user. If you use an application access policy to restrict access to the API, make sure this user is included.
$MsgFrom = "thanos@thejimmartin.com"
# Use the same approach to define CC recipients for the message
#$CcRecipient = "someone@thejimmartin.com"
# Define attachment to send to new users
#$AttachmentFile = "C:\temp\Test.docx"
#$ContentBase64 = [convert]::ToBase64String( [System.IO.File]::ReadAllBytes($AttachmentFile))
$EmailRecipient = "jmartin@contoso.com"
$MsgSubject = "Do you want to see some SMTP OAuth"
$HtmlMsg = "Sample body of the message"
# Create message body and properties and send
$MessageParams = @{
    "URI"         = "https://graph.microsoft.com/v1.0/users/$MsgFrom/sendMail"
    "Headers"     = $Headers
    "Method"      = "POST"
    "ContentType" = 'application/json'
    "Body" = (@{
        "message" = @{
        "subject" = $MsgSubject
        "body"    = @{
        "contentType" = 'HTML' 
        "content"     = $htmlMsg }
        #"attachments" = @(
        #   @{
        #    "@odata.type" = "#microsoft.graph.fileAttachment"
        #    "name" = $AttachmentFile
        #    "contenttype" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        #    "contentBytes" = $ContentBase64 } )  
        "toRecipients" = @(
            @{
            "emailAddress" = @{"address" = $EmailRecipient}
        } ) 
        #"ccRecipients" = @(
        # @{
        #   "emailAddress" = @{"address" = $CcRecipient}
        # }
        }
    }) | ConvertTo-JSON -Depth 6
}
Write-Host "Messages will be sent from $($MsgFrom) to $($EmailRecipient)"
# Send the message
Invoke-RestMethod @MessageParams