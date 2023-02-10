# Install-Module Exchange Online Management - https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps
# install-module ExchangeOnlineManagement

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$ConnectAs,
    [Parameter(Mandatory = $true)]
    [string]$Alias,
    [Parameter(Mandatory = $true)]
    [string]$DisplayName,
    [Parameter(Mandatory = $true)]
    [string]$GivePermissionTo
)

# Connexion Exchange Online
Connect-ExchangeOnline -UserPrincipalName $($ConnectAs + "@csvt.qc.ca")

New-Mailbox -name "$Alias" -Shared -DisplayName "$DisplayName" -Alias "$Alias"
#Set-Mailbox -Identity $Alias -GrantSendOnBehalfTo ferlanda -MessageCopyForSendOnBehalfEnabled:$true

Add-MailboxPermission -Identity $Alias -User $GivePermissionTo -AccessRights FullAccess -InheritanceType All
Set-Mailbox -Identity $Alias -EmailAddresses @{add="$Alias@cssvt.gouv.qc.ca"}
Set-Mailbox -Identity $Alias -EmailAddresses "SMTP:$Alias@cssvt.gouv.qc.ca","smtp:$Alias@csvtqcca.onmicrosoft.com"
get-mailbox -Identity $Alias | FL PrimarySmtpAddress, EmailAddresses
Add-RecipientPermission -Identity $Alias -AccessRights SendAs -Trustee $GivePermissionTo -Confirm:$false
Get-RecipientPermission -Identity $Alias
set-MailboxRegionalConfiguration -Identity $Alias -Language fr-CA -LocalizeDefaultFolderName:$true
Get-MailboxRegionalConfiguration -Identity $Alias