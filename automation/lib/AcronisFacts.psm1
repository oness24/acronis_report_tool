# automation/lib/AcronisFacts.psm1
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "AcronisAnalysis.psm1") -Force
function New-ClientFacts {
    param([hashtable]$Parts)
    # Collect incidents; avoid if-expression empty-array-to-null collapse in PS 5.1
    [object[]]$incidents = @()
    if ($Parts.ContainsKey('incidents') -and $Parts.incidents) {
        [object[]]$incidents = @($Parts.incidents)
    } elseif ($Parts.backup) {
        [object[]]$incidents = @(Get-CriticalIncidents $Parts.backup)
    }
    $facts = [ordered]@{
        meta          = $Parts.meta
        serviceMap    = $Parts.serviceMap
        serviceUsage  = $Parts.serviceUsage
        storage       = $Parts.storage
        backup        = $Parts.backup
        incidents     = $incidents
        security      = $Parts.security
        vulnerability = $Parts.vulnerability
        lifecycle     = $Parts.lifecycle
        audit         = $Parts.audit
        contract      = $Parts.contract
        manual        = $Parts.manual
        outOfContract = @(if ($Parts.ContainsKey('outOfContract') -and $Parts.outOfContract) { $Parts.outOfContract } else { @() })
    }
    # Build a pscustomobject that preserves typed arrays (cast from [ordered] loses @() -> $null on PS 5.1)
    $factsObj = New-Object pscustomobject
    foreach ($k in $facts.Keys) { $factsObj | Add-Member -NotePropertyName $k -NotePropertyValue $facts[$k] }
    $facts.analysis = Get-Analysis $factsObj
    return $facts
}
Export-ModuleMember -Function New-ClientFacts
