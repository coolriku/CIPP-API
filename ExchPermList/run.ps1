using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$APIName = $TriggerMetadata.FunctionName
Write-LogMessage -user $request.headers.'x-ms-client-principal' -API $APINAME  -message "Accessed this API" -Sev "Debug"


# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Define Arrays
$ArrayMailboxPerms = @()
$ArrayMailboxCalFolPerms = @()

# Get Tenant
$TenantFilter = $Request.Query.TenantFilter

#Get Perms
try {
    $mailboxes = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($TenantFilter)/Mailbox" -Tenantid $TenantFilter -scope ExchangeOnline | Where-Object {$_.RecipientTypeDetails -notin @('DiscoveryMailbox')}
    foreach ($Mailbox in $Mailboxes){
        $PermsRequest = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($TenantFilter)/Mailbox('$($Mailbox.PrimarySmtpAddress)')/MailboxPermission" -Tenantid $tenantfilter -scope ExchangeOnline 
        $MailboxPerms = foreach ($Perm in $PermsRequest) {
            if ($Perm.User -notin @('NT AUTHORITY\SELF','Discovery Management')) {
                [pscustomobject]@{
                    Identity = $Mailbox.UserPrincipalName  
                    User         = $Perm.User
                    AccessRights = $Perm.PermissionList.AccessRights -join ', '
                }
            }   
        }
        $ArrayMailboxPerms += $mailboxPerms
    }
    foreach ($Mailbox in $Mailboxes){
        $GetCalParam = @{Identity = $Mailbox.PrimarySmtpAddress; FolderScope = 'Calendar' }
        $CalendarFolder = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-MailboxFolderStatistics" -cmdParams $GetCalParam | Select-Object -First 1
        $CalParam = @{Identity = "$($Mailbox.PrimarySmtpAddress):\$($CalendarFolder.name)" }
        $MailboxCalPermRequest = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-MailboxFolderPermission" -cmdParams $CalParam | Select-Object Identity, User, AccessRights, FolderName
        $MailboxCalPerms = foreach ($CalPerm in $MailboxCalPermRequest) {
            [pscustomobject]@{
                Identity = $Mailbox.UserPrincipalName  
                User         = if ($name = ($Mailboxes | Where-Object {$_.DisplayName -eq $CalPerm.User}).PrimarySmtpAddress) {$name} else {$CalPerm.User}
                AccessRights = $CalPerm.AccessRights -join ', '
                FolderName = $CalPerm.FolderName
            }
        }
        $ArrayMailboxCalFolPerms += $MailboxCalPerms
    }

    $StatusCode = [HttpStatusCode]::OK
    $GraphRequest = [pscustomobject]@{
        MailPerms         = $ArrayMailboxPerms
        CalendarPerms = $ArrayMailboxCalFolPerms
        Mailboxes = $Mailboxes
    }
}
catch {
    $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
    $StatusCode = [HttpStatusCode]::Forbidden
    $GraphRequest = $ErrorMessage
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = $StatusCode
        Body       = @($GraphRequest)
    })
