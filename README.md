HOW TO USE

Create a variable with the list of possible recipients

$TjmRecipients = (Get-Recipient -ResultSize unlimited | Where-Object {$_.PrimarySmtpAddress -notlike "*onmicrosoft*"}).PrimarySmtpAddress

Create a variable with the list of possible senders

$TjmMailboxes = (Get-Mailbox -ResultSize unlimited | Where-Object {$_.PrimarySmtpAddress -notlike "*onmicrosoft*"}).PrimarySmtpAddress

Create a secure string variable for the app secret

$secret = ConvertTo-SecureString -String "1MW8Q~MxJHRYYUzn21QbIialJMGjG8Nzz-QUUaUA" -AsPlainText -Force


Syntax to send 10 messages from random senders to random recipients

.\Graph-SendMessages.ps1 -OAuthClientId 6a93c8c4-9cf6-4efe-a8ab-9eb178b8dff4 -OAuthTenantId 9101fc97-5be5-5538-a1d7-83e051e52057 -OAuthClientSecret $secret -PermissionType Application -ToRecipients $TjmRecipients -Sender $TjmMailboxes -NumberOfMessages 10


Syntax to send 5 messages from a single sender to two recipients with an attachment in each message

.\Graph-SendMessages.ps1 -OAuthClientId 6a93c8c4-9cf6-4efe-a8ab-9eb178b8dff4 -OAuthTenantId 9101fc97-5be5-5538-a1d7-83e051e52057 -OAuthClientSecret $secret -PermissionType Application -ToRecipients user1@contoso.com,user2@contoso.com -RandomRecipients:$false -Sender sender@contoso.com -NumberOfMessages 5 AttachmentPath C:\Scripts\Attachments\


PARAMETERS

The AzureEnvironment parameter specifies the Azure environment for the tenant.

The PermissionType parameter specifies whether the app registrations uses delegated or application permissions.
    
The OAuthClientId parameter is the Azure Application Id that this script uses to obtain the OAuth token.  Must be registered in Azure AD.
    
The OAuthTenantId parameter is the tenant Id where the application is registered (Must be in the same tenant as mailbox being accessed).

The OAuthRedirectUri parameter is the redirect Uri of the Azure registered application.
    
The OAuthSecretKey parameter is the the secret for the registered application.
    
The OAuthCertificate parameter is the certificate for the registered application.
  
The CertificateStore parameter specifies the certificate store where the certificate is loaded.

The Scope parameter specifies the Graph API permission needed to run the script.

The NumberOfMessages parameter specifies the number of messages the script should attempt to send.

The Sender parameter specifies one or more email addresses to be used as the sender of the message.

The ToRecipients parameter specifies one or more email addresses to be used in the To field of the message.

The CcRecipients parameter specifies one or more email addresses to be used in the Cc field of the message.

The BccRecipients parameter specifies one or more email addresses to be used in the Bcc field of the message.

The AttachmentPath parameter specifies the directory where attachments can be found.

The RandomRecipients parameter specifies whether the recipients should be randomly selected from available email addresses.
