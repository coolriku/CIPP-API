using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

$APIName = $TriggerMetadata.FunctionName
Write-LogMessage -user $request.headers.'x-ms-client-principal' -API $APINAME -message 'Accessed this API' -Sev 'Debug'


# Write to the Azure Functions log stream.
Write-Host 'PowerShell HTTP trigger function processed a request.'

# Interact with query parameters or the body of the request.
$TenantFilter = $Request.Query.TenantFilter



Try{
    $Table = Get-CIPPTable -TableName cacheexpermrpt
    #Checking for Loading table entry, so no new report is queued.
    $Loading = Get-AzDataTableEntity @Table | Where-Object {$_.Timestamp -GT (Get-Date).AddMinutes(-30) -and ($_.Tenant -eq $TenantFilter) -and ($_.Report -eq 'Loading')}
    If ($loading.Report -eq 'Loading'){
        $GraphRequest = [pscustomobject]@{
            Identity = $null
            User         = 'Already Loading. Please be more patient'
            AccessRights = $null
            Type = $null
            FolderName = $null
        }
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::OK
            Body       = @($GraphRequest)
        })
        exit
    }
    #Else Try to load Report from table.
    $Rows = Get-AzDataTableEntity @Table | Where-Object {$_.Timestamp -GT (Get-Date).AddMinutes(-30) -and ($_.Tenant -eq $TenantFilter)}
    #If no rows, than queue new report creation
    if (!$Rows) {
        Push-OutputBinding -Name Msg -Value $TenantFilter
        $GraphRequest = [pscustomobject]@{
            Identity = $null
            User         = 'Loading data. Please check back in 1 minute'
            AccessRights = $null
            Type = $null
            FolderName = $null
        }
    }
    #Load Report
    else {
        $GraphRequest = $Rows.Report | ConvertFrom-Json
    }
}
catch{
    $ErrorMessage = Get-NormalizedError -Message $_.Exception.Message
    $GraphRequest = [pscustomobject]@{
        Identity        = $null
        User            = $ErrorMessage
        AccessRights    = $null
        Type            = $null
        FolderName      = $null
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