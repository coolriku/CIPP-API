using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$APIName = $TriggerMetadata.FunctionName
Write-LogMessage -user $request.headers.'x-ms-client-principal' -API $APINAME -message 'Accessed this API' -Sev 'Debug'


# Write to the Azure Functions log stream.
Write-Host 'PowerShell HTTP trigger function processed a request.'

# Interact with query parameters or the body of the request.
$TenantFilter = $Request.Query.TenantFilter

$Table = Get-CIPPTable -TableName cacheexpermrpt
# Remove-AzDataTableEntity @Table -Entry (Get-AzDataTableEntity @Table)
# exit
$Loading = Get-AzDataTableEntity @Table | Where-Object {$_.Timestamp -GT (Get-Date).AddMinutes(-30) -and ($_.Tenant -eq $TenantFilter) -and ($_.Report -eq 'Loading')}
If ($loading.Report -eq 'Loading'){
    $GraphRequest = [pscustomobject]@{
        Tenant      = $null
        Timestamp   = $null
        Report      = $GraphRequest = [pscustomobject]@{
            Identity = 'Already Loading. Please be more patient'
            User         = $null
            AccessRights = $null
            Type = $null
            FolderName = $null
        }
    }
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @($GraphRequest)
    })
    exit
}
$Rows = Get-AzDataTableEntity @Table | Where-Object {$_.Timestamp -GT (Get-Date).AddMinutes(-30) -and ($_.Tenant -eq $TenantFilter)}
if (!$Rows) {
    Push-OutputBinding -Name Msg -Value $TenantFilter
    $GraphRequest = [pscustomobject]@{
        Tenant      = $_.Tenant
        Timestamp   = $_.Timestamp
        Report      = [pscustomobject]@{
            Identity = 'Loading data. Please check back in 1 minute'
            User         = $null
            AccessRights = $null
            Type = $null
            FolderName = $null
        }
    }
}         
else {
    $GraphRequest = $Rows | Where-Object -Property Tenant -EQ $TenantFilter | ForEach-Object { 
        [pscustomobject]@{
            Tenant      = $_.Tenant
            Timestamp   = $_.Timestamp
            Report      = $_.Report | ConvertFrom-Json
        }
    }
}
#Remove all old cache
try{
    Remove-AzDataTableEntity @Table -Entity (Get-AzDataTableEntity @Table | Where-Object {$_.Timestamp -LT (Get-Date).AddMinutes(-30) -and ($_.Tenant -eq $TenantFilter)})
}
catch {
    write-host "error removing datatable entry"
}
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @($GraphRequest)
    })