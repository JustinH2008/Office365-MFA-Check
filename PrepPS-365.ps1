Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NetFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Value '1' -Type DWord
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\.NetFramework\v4.0.30319' -Name 'SchUseStrongCrypto' -Value '1' -Type DWord
Set-ExecutionPolicy Unrestricted
Install-Module PowershellGet -Force
Install-Module -Name CredentialManager
Install-Module -Name ExchangeOnlineManagement
Install-Module -Name AzureAD
Install-Module -Name MSOnline
