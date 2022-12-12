using namespace System.Net

# Input bindings are passed in via param block.
#param($Request, $TriggerMetadata)
param([string] $QueueItem, $TriggerMetadata)

$APIName = $TriggerMetadata.FunctionName
Write-LogMessage -user $request.headers.'x-ms-client-principal' -API $APINAME  -message "Accessed this API" -Sev "Debug"


# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

#Define Arrays
$ArrayPermissions = @()
# Get Tenant
$TenantFilter = $QueueItem

#create entry that report is loading.
$Table = Get-CIPPTable -TableName cacheexpermrpt
$LoadingGuid = (New-Guid).guid
$RowLoading = @{
    RowKey  = [string]$LoadingGuid
    Tenant = [string]$TenantFilter
    Report = 'Loading'
    PartitionKey = 'ExchPermReport'
}
Add-AzDataTableEntity @Table -Entity $RowLoading -Force | Out-Null

try {
    #Get Mailboxes
    $mailboxes = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($TenantFilter)/Mailbox" -Tenantid $TenantFilter -scope ExchangeOnline | Where-Object {$_.RecipientTypeDetails -notin @('DiscoveryMailbox')}
    foreach ($Mailbox in $Mailboxes){
        #Get Mailbox Permissions
        $PermsRequest = New-GraphGetRequest -uri "https://outlook.office365.com/adminapi/beta/$($TenantFilter)/Mailbox('$($Mailbox.PrimarySmtpAddress)')/MailboxPermission" -Tenantid $tenantfilter -scope ExchangeOnline 
        $MailboxPerms = foreach ($Perm in $PermsRequest) {
            if ($Perm.User -notin @('NT AUTHORITY\SELF','Discovery Management')) {
                [pscustomobject]@{
                    Identity = $Mailbox.UserPrincipalName  
                    User         = $Perm.User
                    AccessRights = $Perm.PermissionList.AccessRights -join ', '
                    Type = 'Mailbox'
                    FolderName = $null
                }
            }   
        }
        $ArrayPermissions += $mailboxPerms
    }
    #Get Send As / On Behalf Permissions
    $RecipPerms = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-RecipientPermission" -cmdParams @{ ResultSize = 'Unlimited' } | Where-Object {$_.Identity -in $mailboxes.Identity}
    foreach ($RecipPerm in $RecipPerms){
        if($RecipPerm.TrusteeSidString -eq 'S-1-5-10'){
            #do nothing
        }
        else{
            $ArrayPermissions +=
            [pscustomobject]@{
                Identity = ($Mailboxes | Where-Object {$_.Identity -eq $RecipPerm.Identity}).PrimarySmtpAddress
                User         = $RecipPerm.Trustee
                AccessRights = $RecipPerm.AccessRights -join ', '
                Type = 'Mailbox'
                FolderName = $null
            }
        }
    }
    #GCalendar Permissions Loop
    foreach ($Mailbox in $Mailboxes){
        #Get Calendar Folders
        $GetCalParam = @{Identity = $Mailbox.PrimarySmtpAddress; FolderScope = 'Calendar' }
        $CalendarFolders = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-MailboxFolderStatistics" -cmdParams $GetCalParam
        #Get Root Calendar Folder
        $RootCalendarFol = $CalendarFolders | Where-Object {$_.FolderType -eq 'Calendar'}
        #Loop through Calendar Folders except default birthday
        Foreach ($CalendarFolder in $CalendarFolders | where-object {$_.ContainerClass -ne 'IPF.Appointment.Birthday'}){
            if ($CalendarFolder.FolderType -eq 'Calendar'){
                $CalParam = @{Identity = "$($Mailbox.PrimarySmtpAddress):\$($CalendarFolder.name)" }
            }
            else {
                $CalParam = @{Identity = "$($Mailbox.PrimarySmtpAddress):\$($RootCalendarFol.name)\$($CalendarFolder.name)" }
            }
            #Get Calendar Folder Permissions
            $MailboxCalPermRequest = New-ExoRequest -tenantid $TenantFilter -cmdlet "Get-MailboxFolderPermission" -cmdParams $CalParam
            $MailboxCalPerms = foreach ($CalPerm in $MailboxCalPermRequest) {
                if(($CalPerm.User -in @('Default','Anonymous')) -and ($CalPerm.AccessRights[0] -eq 'None')){
                    #do nothing
                }
                else {
                    [pscustomobject]@{
                        Identity = $Mailbox.UserPrincipalName  
                        User         = if ($name = ($Mailboxes | Where-Object {$_.DisplayName -eq $CalPerm.User}).PrimarySmtpAddress) {$name} else {$CalPerm.User}
                        AccessRights = $CalPerm.AccessRights -join ', '
                        Type = 'Calendar'
                        FolderName = $CalPerm.FolderName
                    }
                }
            }
            $ArrayPermissions += $MailboxCalPerms
        }
    }

    $GraphRequest = $ArrayPermissions
}
catch {
    $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
    $GraphRequest = [pscustomobject]@{
        Identity = [string]$ErrorMessage
        User         = [string]$null
        AccessRights = [string]$null
        Type = [string]$null
        FolderName = [string]$null
    }
}

$Row = @{
    RowKey  = [string](New-Guid).guid
    Tenant = [string]$TenantFilter
    Report = [string]( $GraphRequest | ConvertTo-Json -Compress )
    PartitionKey = 'ExchPermReport'
}

Remove-AzDataTableEntity @Table -Entry (Get-AzDataTableEntity @Table | Where-Object {$_.RowKey -eq $LoadingGuid})
Add-AzDataTableEntity @Table -Entity $Row -Force | Out-Null