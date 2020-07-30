using namespace System.Net

# Set a parameter
param($Request)

# Instantiate Credentials
$username = $Env:user
$pw = $Env:password

# Build Credentials
$keypath = "D:\home\site\wwwroot\GetO365User\bin\keys\PassEncryptKey.key"
$secpassword = $pw | ConvertTo-SecureString -Key (Get-Content $keypath)
$credential = New-Object System.Management.Automation.PSCredential ($username, $secpassword)


# Import the MSOnline Module and establish a session to Exchange Online
Import-Module MSOnline -UseWindowsPowerShell


#Import the MSOnline Module and establish a session to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

#Create a variable to pass a value from the Request to Powershell

$identity = $Request.Query.Identity

# add User from O365 GAL
$result = Set-MailUser -Identity $identity -HiddenFromAddressListsEnabled $false

#Compose an HTTP Response to send when the command is run 

$HttpResponse = [HttpResponseContext]@{ 
 StatusCode = 200
 Body = $result
} 

#Send the Response back 

Push-OutputBinding -Name Response -Value $HttpResponse

#Terminate the Exchange Session
Remove-PSSession $Session
 
