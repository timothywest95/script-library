<#
.SYNOPSIS
   Sends an email with a customizable message and optional attachments.
.DESCRIPTION
   This script sends an email via SMTP, using the Send-MailKitMessage (https://github.com/austineric/Send-MailKitMessage). 
   It supports authenticated SMTP and can store credentials securely for re-use. It will prompt for the SMTP user's password on first run and save to a file for re-use.
   Tested using SMTP2GO.
.NOTES
  Version:        1.0
  Author:         Timothy West
  Creation Date:  2024-08-10
  Sources: Thanks to the developers of Send-MailKitMessage (https://github.com/austineric/Send-MailKitMessage). This page also helped: https://www.alitajran.com/send-email-powershell/
#>

# Change these variables

$smtpUsername = "sender@mydomain.com"
$fromAddress = "sender@mydomain.com" # Usually the same as $smtpUsername
$passwordFilePath = "C:\scripts\smtp_creds.txt" # Path to the file where the encrypted password will be stored
$SMTPServer = "mail.smtp2go.com" # SMTP server ([string], required)
$Port = "2525" # Port ([int], required)
$recipientAddress = "jdoe@contoso.com"
#$CCRecipientAddress = "rroe@contoso.com"
#BCCRecipientAddress = "jsmith@fabrikam.com"
$emailSubject = "Test 1"
#$emailTextBody = "Hello, world."
$emailHTMLBody = @"
<h1>Test 1</h1>
<p>Hello, world.</p>
"@

# Import Module
Import-Module Send-MailKitMessage

# Check if the password file exists
if (-Not (Test-Path -Path $passwordFilePath)) {
    # Prompt the user to enter the password and save it securely to a file
    $securePassword = Read-Host -Prompt "SMTP password not found. Enter SMTP password. This will be stored for future use." -AsSecureString
    $securePassword | ConvertFrom-SecureString | Out-File -FilePath $passwordFilePath
    Write-Host "Password has been saved securely. The stored password will be used for future script runs."
} else {
    # Read and decrypt the password from the file
    $securePassword = Get-Content -Path $passwordFilePath | ConvertTo-SecureString
}

# Authentication ([System.Management.Automation.PSCredential], optional)
$Credential = [System.Management.Automation.PSCredential]::new("$smtpUsername", $securePassword)

# Sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From = [MimeKit.MailboxAddress]"$fromAddress"

# Recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList = [MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]"$recipientAddress")

# CC list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$CCList = [MimeKit.InternetAddressList]::new()
#$CCList.Add([MimeKit.InternetAddress]"$CCRecipientAddress")

# BCC list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
#$BCCList = [MimeKit.InternetAddressList]::new()
#$BCCList.Add([MimeKit.InternetAddress]"$BCCRecipientAddress")

# Subject ([string], optional)
$Subject = [string]"$emailSubject"

# Text body ([string], optional)
#$TextBody = [string]"$emailTextBody"

# HTML body ([string], optional)
$HTMLBody = [string]"$emailHTMLBody"

# Attachment list ([System.Collections.Generic.List[string]], optional)
#$AttachmentList = [System.Collections.Generic.List[string]]::new()
#$AttachmentList.Add("Attachment1FilePath")

# Splat parameters
$Parameters = @{
    "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable    
    "Credential"                     = $Credential
    "SMTPServer"                     = $SMTPServer
    "Port"                           = $Port
    "From"                           = $From
    "RecipientList"                  = $RecipientList
    "CCList"                         = $CCList
    "BCCList"                        = $BCCList
    "Subject"                        = $Subject
    "TextBody"                       = $TextBody
    "HTMLBody"                       = $HTMLBody
    "AttachmentList"                 = $AttachmentList
}

# Send message
Send-MailKitMessage @Parameters
