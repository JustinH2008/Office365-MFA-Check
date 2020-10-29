#region Grab Credentials from CredManager and connect to 365 via PS

#import credential manager module for secure password
   Import-Module CredentialManager

#Variables
$target = Get-StoredCredential -Target 
# ^ This is what you set in credential manager for 365 access creds.
$smtpServer = 'InsertSMTPserver here'
#^ This is the SMTP server that you are sending from.
$sentFrom = 'SentFrom@service.com'
#^ The email that this is being sent from.
$sendTo = 'Myemail@email.com'
#^ Who this is being sent to
$subject = 'MFA Compliance Check Data'
#^ Subject of the email being sent
$body = '!!Contact:'+ $SentFrom +'!!
Attached is the Office 365 user list with MFA status.'

#^ Body of the email being sent.

#Create remote Powershell session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential ($target) -Authentication Basic –AllowRedirection    

#Import the session
   Import-PSSession $Session -AllowClobber | Out-Null

   #Connect to Azure AD
   Import-Module MSOnline
   Connect-MsolService -Credential ($target)

   ### Azure AD v2.0
   Connect-AzureAD -Credential ($target)

#endregion


#region Initialize Variables

$MemberObjects = @()
$OutputFile = "c:\MFA Compliance Check\MFAStatus.csv"

#endregion

#region Pull 365info and export

$Users = Get-MsolUser -EnabledFilter EnabledOnly -All | where {$_.UserPrincipalName -notlike "*onmicrosoft.com"} | select UserPrincipalName,Department,StrongAuthenticationRequirements
foreach ($User in $Users) {
    $MemberObject = New-Object System.Object
    $MemberObject | Add-Member -MemberType NoteProperty -Name "UPN" -Value $User.UserPrincipalName
    $MemberObject | Add-Member -MemberType NoteProperty -Name "Location" -Value $User.Department
    $MemberObject | Add-Member -MemberType NoteProperty -Name "MFAState" -Value $User.StrongAuthenticationRequirements.State
    $MemberObjects += $MemberObject
    }

$MemberObjects | Export-Csv -Path $OutputFile -NoTypeInformation

#endregion


#region Send email to Team

Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue
$file = "C:\MFA Compliance Check\MFAStatus.csv"

$mailboxdata = (Get-MailboxStatistics | select DisplayName, TotalItemSize,TotalDeletedItemSize, ItemCount, LastLoggedOnUserAccount, LastLogonTime)

$att = new-object Net.Mail.Attachment($file)

$msg = new-object Net.Mail.MailMessage

$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$msg.From = $sentFrom

$msg.To.Add($sendTo)

$msg.Subject = $subject

$msg.Body = $body

$msg.Attachments.Add($att)

$smtp.Send($msg)

$att.Dispose()

#endregion


#Clean Up
cd 'C:\MFA Compliance Check'
Remove-Item MFAStatus.csv