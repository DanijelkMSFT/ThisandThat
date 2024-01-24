
function Get-QuarantineEntries {
    $Page = 1
    $List = @()
    $PageList = @()
    Do
    {
        Write-Host ("- Getting Page {0}" -f $Page)
        $PageList = (Get-QuarantineMessage -Type HighConfPhish -PageSize 1000 -Page $Page | Where-Object {$_.Released -like "False" -and $_.SenderAddress -like "*"})
        Write-Host ("-- {0} rows in this page match" -f $PageList.count)
        $List += $PageList
        Write-Host "--- Exporting list to CSV for logging"
        $Pagelist | Export-Csv -Path "C:\temp\Quarantined Message Matches.csv" -Append -NoTypeInformation
        $Page = $Page + 1
    } while ($PageList.count -gt 0)
    return $List
}

function Release-QuarantineEntries {
    param (
		[Parameter(Mandatory = $true)][PSCustomObject]$MessageList,
        [Parameter(Mandatory = $true)][int]$batchSize
        )

    $batchCount = [math]::Ceiling($MessageList.length / $batchSize)
    
    for ($i = 0; $i -lt $batchCount; $i++) {
        $batch = $MessageList.Identity[($i * $batchSize)..(($i + 1) * $batchSize - 1)]
        Write-Host "Processing batch $($i + 1): $batch"
        Write-Host("The following amount {0} of messages are going to be released" -f $batch.count)
        Release-QuarantineMessage -identities $batch -ReleaseToAll -WhatIf
    }
}

## create List of HighConfPhish and save it to c:\temp\Quarantined Message Matches.csv
$QList = Get-QuarantineEntries
$QList = $null

# load csv 
$Qlist = Import-Csv -LiteralPath 'c:\temp\Quarantined Message Matches.csv'

## Process List in batches (max. 100 supported)
Release-QuarantineEntries -MessageList $QList -batchSize 100
