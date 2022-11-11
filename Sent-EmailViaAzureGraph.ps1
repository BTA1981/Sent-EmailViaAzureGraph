# Currently (11-2022), Azure Graph only works with USER certificates or an access token.
# This test script works by authenticating via token (if you want to use a password) and can be changed to authenticate via Certificate


$AppID = "0a8e4424-56ea-42b4-8926-3a09fad36cbb" # App ID NK_SendEmailScript App registration
$tenantID = "e71710f3-ac90-4283-966f-0b9be5d0cd17"
#$thumbprint = ""
$O365Organisation = "novadickentron.onmicrosoft.com"
$CredPath = "C:\Beheer\Scripts\New-EverWorkAccount\NK\Credentials\NKazureGraph_.xml"
$KeyFilePath = "C:\Beheer\Scripts\New-EverWorkAccount\NK\Credentials\NKazureGraph_.key"

#region Decrypt password
$Key = Get-Content $KeyFilePath
$credXML = Import-Clixml $CredPath # Import encrypted credential file into XML format
$secureStringPWD = ConvertTo-SecureString -String $credXML.Password -Key $key
$Credentials = New-Object System.Management.Automation.PsCredential($credXML.UserName, $secureStringPWD) # Create PScredential Object
$passwordplain = (New-Object PSCredential 0, $secureStringPWD).GetNetworkCredential().Password
#endregion

#region needed to create access token
$body =  @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $AppID
    Client_Secret = $passwordplain
}
 
$connection = Invoke-RestMethod `
    -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
    -Method POST `
    -Body $body
 
$token = $connection.access_token
#endregion

Connect-MgGraph -AccessToken $token

# Need to get certificate
#$cert = Get-ChildItem "Cert:\LocalMachine\My\$($thumbprint)"

#Connect-MgGraph -TenantId $tenantID -ClientId $AppID -CertificateThumbprint $thumbprint
#Connect-MgGraph -TenantId $tenantID - -CertificateThumbprint $thumbprint -ClientId $AppID

[string]$EmailTemplate1 = 'C:\Beheer\NewUserEmailHTML.html'	
	
$user = "client@novadic-kentron.nl"
$name = "Eshgro"
$company = "test"
$subject = "test mail"
$type = "html"
 
$template = Get-Content `
    -path "$EmailTemplate1" `
    -raw
 
$template = $template.Replace('{{NAME}}',$name)
$template = $template.Replace('{{COMPANY}}',$company)
$content = $template
 
$recipients = @()
$recipients += @{
    emailAddress = @{
        address = $user
    }
}
$message = @{
    subject = $subject;
    toRecipients = $recipients;
    body = @{
        contentType = $type;
        content = $content
    }
}
 
Send-MgUserMail `
    -UserId $user `
    -Message $message

