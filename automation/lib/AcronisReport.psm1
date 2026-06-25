# AcronisReport.psm1 - read-only collection + aggregation helpers for MSSP monthly reports.
# All functions are GET-only against the Acronis Cyber Protect Cloud API.
# Conventions baked in (see project CLAUDE.md):
#   - Storage shown in BINARY units to match the Acronis console (bytes / 1GB).
#   - task_manager activities need a TENANT-SCOPED token; cursor pages carry ONLY `after`.
#   - agent_manager needs a tenant-scoped token and NO params.
#   - PowerShell 5.1 hides HTTP error bodies, so activities use the curl binary.
# Cross-platform: Windows PowerShell 5.1 and pwsh 7 on Linux/macOS (Google Colab).

$ErrorActionPreference = "Stop"

# Decode child-process (curl) stdout as UTF-8 so accented names (Serviço, Água) aren't mojibake on PS 5.1.
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# 'curl.exe' on Windows (avoids the PS curl->Invoke-WebRequest alias); 'curl' on Linux/macOS.
$script:CurlBin = if ($IsLinux -or $IsMacOS) { 'curl' } else { 'curl.exe' }

function Get-AcronisToken {
    <# Mint an OAuth2 client-credentials token. Pass -TenantId to scope it to one customer. #>
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$Secret,
        [string]$TenantId
    )
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${ClientId}:${Secret}"))
    $body = @{ grant_type = "client_credentials" }
    if ($TenantId) { $body.scope = "urn:acronis.com:tenant-id:$TenantId" }
    $r = Invoke-RestMethod -Method Post -Uri "$BaseUrl/api/2/idp/token" `
        -Headers @{ Authorization = "Basic $b64" } -Body $body `
        -ContentType "application/x-www-form-urlencoded"
    return $r.access_token
}

function Get-AcronisCustomers {
    <# List customer tenants directly under a partner root. #>
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [Parameter(Mandatory)][string]$PartnerToken,
        [Parameter(Mandatory)][string]$RootTenantId
    )
    # Use curl (UTF-8 stdout) instead of Invoke-RestMethod, which mis-decodes accents on PS 5.1.
    $raw = Invoke-Curl -Token $PartnerToken -Url "$BaseUrl/api/2/tenants?parent_id=$RootTenantId" -MaxTime 90
    return ($raw | ConvertFrom-Json).items | Where-Object { $_.kind -eq "customer" }
}

function Invoke-Curl {
    <# Thin wrapper so the curl call site is identical everywhere and OS-correct. #>
    param([string]$Token, [string]$Url, [int]$MaxTime = 90)
    return & $script:CurlBin -s --max-time $MaxTime -H "Authorization: Bearer $Token" $Url
}

function Get-AcronisUsages {
    <# Raw /usages items for a tenant (partner token works). #>
    param([string]$BaseUrl, [string]$PartnerToken, [string]$TenantId)
    $raw = Invoke-Curl -Token $PartnerToken -Url "$BaseUrl/api/2/tenants/$TenantId/usages" -MaxTime 90
    return ($raw | ConvertFrom-Json).items
}

function Get-ServiceProfile {
    <# Derive a service profile from /usages: enabled offering items (status=1) and their use. #>
    param($UsageItems)
    $enabled = $UsageItems | Where-Object { $_.offering_item.status -eq 1 }
    $get = { param($n) ($UsageItems | Where-Object { $_.usage_name -eq $n } | Measure-Object value -Sum).Sum }

    $servers      = [int](& $get "servers")
    $vms          = [int](& $get "vms")
    $workstations = [int](& $get "workstations")
    $advSecurity  = [int](& $get "adv_security_workloads")
    $cloudBytes   = [double](($UsageItems | Where-Object { $_.usage_name -eq "storage" } | Measure-Object value -Sum).Sum)
    $localBytes   = [double](($UsageItems | Where-Object { $_.usage_name -eq "local_storage_total" } | Measure-Object value -Sum).Sum)

    [pscustomobject]@{
        hasBackup          = [bool]($enabled | Where-Object { $_.usage_name -in @("servers","vms","workstations") })
        hasEDR             = [bool]($enabled | Where-Object { $_.usage_name -match "adv_security|edr" })
        servers            = $servers
        vms                = $vms
        workstations       = $workstations
        protectedWorkloads = $servers + $vms + $workstations
        advSecurity        = $advSecurity
        cloudStorageBytes  = $cloudBytes
        cloudStorageGB     = [math]::Round($cloudBytes / 1GB, 2)
        localStorageGB     = [math]::Round($localBytes / 1GB, 2)
        enabledItems       = ($enabled | Select-Object -ExpandProperty usage_name -Unique)
    }
}

function Get-MonthBounds {
    <# Returns UTC ISO bounds for a YYYY-MM month string. #>
    param([Parameter(Mandatory)][string]$Month)
    $start = [datetime]::ParseExact("$Month-01", "yyyy-MM-dd", $null)
    $end   = $start.AddMonths(1)
    [pscustomobject]@{
        StartIso = $start.ToString("yyyy-MM-ddT00:00:00Z")
        EndIso   = $end.ToString("yyyy-MM-ddT00:00:00Z")
    }
}

function Get-BackupActivities {
    <# Paginated month backup activities with a tenant-scoped token. Cursor pages carry ONLY `after`. #>
    param(
        [string]$BaseUrl, [string]$TenantToken,
        [string]$StartIso, [string]$EndIso, [int]$MaxPages = 200
    )
    $all = @(); $cursor = $null; $page = 0
    do {
        if ($cursor) {
            $u = "$BaseUrl/api/task_manager/v2/activities?after=$([uri]::EscapeDataString($cursor))"
        } else {
            $u = "$BaseUrl/api/task_manager/v2/activities?policyType=backup&completedAt=gt($StartIso)&completedAt=lt($EndIso)&limit=200"
        }
        $raw = Invoke-Curl -Token $TenantToken -Url $u -MaxTime 280
        if (-not $raw) { Write-Warning "  activities page $page timed out"; break }
        $j = $raw | ConvertFrom-Json
        if ($j.code -or $j.domain) { Write-Warning "  activities API error: $raw"; break }
        $all += $j.items; $page++
        $cursor = $null
        if ($j.paging.cursors.after) { $cursor = $j.paging.cursors.after }
    } while ($cursor -and $page -lt $MaxPages)
    return $all
}

function Get-PagedItems {
    <# Follow alert_manager / agent_manager cursor pagination (paging.cursors.after).
       These endpoints cap a page at 100 (agents) / 200 (alerts); a single call silently
       truncates, so always page through every cursor. Returns the full items array. #>
    param(
        [string]$Token, [string]$Url,
        [int]$MaxPages = 100, [int]$MaxTime = 120
    )
    $all = @(); $cursor = $null; $page = 0
    $sep = if ($Url -match '\?') { '&' } else { '?' }
    do {
        $u = if ($cursor) { "$Url$sep" + "after=$([uri]::EscapeDataString($cursor))" } else { $Url }
        $raw = Invoke-Curl -Token $Token -Url $u -MaxTime $MaxTime
        if (-not $raw) { Write-Warning "  paged fetch timed out (page $page)"; break }
        $j = $raw | ConvertFrom-Json
        if ($j.code -or $j.domain) { Write-Warning "  paged API error: $raw"; break }
        $all += $j.items; $page++
        $cursor = $null
        if ($j.paging.cursors.after) { $cursor = $j.paging.cursors.after }
    } while ($cursor -and $page -lt $MaxPages)
    return ,$all
}

function Measure-Backup {
    <# Enriched backup aggregation: per-machine breakdown, runMode, PT-BR error reasons, successPctExclWorst. #>
    param([object[]]$Activities)
    $roots = @($Activities | Where-Object { $_.context.IsProcessRoot -eq $true })
    if (-not $roots) { $roots = @($Activities) }
    $ok   = @($roots | Where-Object { $_.result.code -eq "ok" }).Count
    $err  = @($roots | Where-Object { $_.result.code -eq "error" }).Count
    $warn = @($roots | Where-Object { $_.result.code -eq "warning" }).Count
    $runs = $roots.Count
    $sched = @($roots | Where-Object { $_.context.runMode -eq "scheduled" }).Count
    $manual = @($roots | Where-Object { $_.context.runMode -eq "manual" }).Count
    $bytes = ($Activities | ForEach-Object { [double]$_.progress.bytesSaved } | Measure-Object -Sum).Sum
    if (-not $bytes) { $bytes = 0 }

    $perMachine = $roots | Group-Object { $_.context.MachineName } | ForEach-Object {
        $m = $_.Group
        [pscustomobject]@{
            machine = $_.Name
            runs    = $_.Count
            ok      = @($m | Where-Object { $_.result.code -eq "ok" }).Count
            err     = @($m | Where-Object { $_.result.code -eq "error" }).Count
        }
    }
    $errorReasons = $roots | Where-Object { $_.result.error.code } |
        Group-Object { $_.result.error.code } | ForEach-Object {
            [pscustomobject]@{ code=$_.Name; plainPt=(Get-ErrorDescription $_.Name); count=$_.Count }
        }
    # success excluding the single worst-failing machine
    $worst = $perMachine | Sort-Object err -Descending | Select-Object -First 1
    $exclRuns = $runs; $exclOk = $ok
    if ($worst -and $worst.err -gt 0) { $exclRuns -= $worst.runs; $exclOk -= $worst.ok }
    [pscustomobject]@{
        totalActivities    = @($Activities).Count
        runs               = $runs
        scheduled          = $sched
        manual             = $manual
        ok                 = $ok
        warning            = $warn
        error              = $err
        successPct         = if ($runs) { [math]::Round(100.0*$ok/$runs,1) } else { 0 }
        successPctExclWorst= if ($exclRuns -gt 0)  { [math]::Round(100.0*$exclOk/$exclRuns,1) }
                              elseif ($runs)        { [math]::Round(100.0*$ok/$runs,1) }
                              else                  { 0 }
        transferredGB      = [math]::Round([double]$bytes/1GB,2)
        perMachine         = @($perMachine)
        errorReasons       = @($errorReasons)
        worstMachine       = if ($worst) { $worst.machine } else { $null }
    }
}

function Measure-Agents {
    <# Agent inventory aggregation: os/version breakdown, online, onboarded-this-month, offline, outdated. #>
    param([object[]]$Agents, [string]$Month)
    $a = @($Agents)
    # platform can be an object {family,name,...} or a plain string; normalise to a label
    $byOs = $a | Group-Object {
        $p = $_.platform
        if ($p -is [string]) { $p } elseif ($p) { if ($p.name) { $p.name } else { $p.family } } else { "unknown" }
    } | ForEach-Object { [pscustomobject]@{ os=$_.Name; count=$_.Count } }
    # core_version can be an object {current:{release_id,...}} or a plain string
    $getVer = { param($agent)
        $cv = $agent.core_version
        if ($cv -is [string]) { $cv }
        elseif ($cv -and $cv.current -and $cv.current.release_id) { $cv.current.release_id }
        else { $null }
    }
    $ver = $a | Group-Object { & $getVer $_ } | ForEach-Object { [pscustomobject]@{ version=$_.Name; count=$_.Count } }
    $newest = ($a | ForEach-Object { & $getVer $_ } | Where-Object { $_ } |
        Sort-Object -Descending { $v=$null; if ([version]::TryParse($_, [ref]$v)) { $v } else { [version]"0.0" } } |
        Select-Object -First 1)
    # registration_date can be RFC-2822 "Fri, 19 Jun 2026 04:43:19 +0000" or ISO; match month prefix in "Jun 2026" or "2026-06"
    $yr = $Month.Split('-')[0]; $mo = $Month.Split('-')[1]
    $monthNames = @{
        '01'='Jan';'02'='Feb';'03'='Mar';'04'='Apr';'05'='May';'06'='Jun';
        '07'='Jul';'08'='Aug';'09'='Sep';'10'='Oct';'11'='Nov';'12'='Dec'
    }
    $moName = $monthNames[$mo]
    $onboard = @($a | Where-Object {
        $rd = "$($_.registration_date)"
        if (-not $rd) { return $false }
        # ISO prefix match
        if ($rd.StartsWith($Month)) { return $true }
        # RFC-2822: contains "May 2026" pattern
        $rd -match "$moName $yr"
    } | ForEach-Object { if ($_.name) { $_.name } else { $_.hostname } })
    $offline = @($a | Where-Object { -not $_.online } | ForEach-Object { if ($_.name) { $_.name } else { $_.hostname } })
    $outdated = @($a | Where-Object {
        $v = & $getVer $_
        $v -and $v -ne $newest
    } | ForEach-Object { if ($_.name) { $_.name } else { $_.hostname } })
    $agentsList = @($a | ForEach-Object {
        [pscustomobject]@{
            name        = $(if ($_.name) { $_.name } else { $_.hostname })
            os          = $(
                $p = $_.platform
                if ($p -is [string]) { $p } elseif ($p) { if ($p.name) { $p.name } else { $p.family } } else { "unknown" }
            )
            coreVersion = $(& $getVer $_)
            online      = [bool]$_.online
        }
    })
    [pscustomobject]@{
        total              = $a.Count
        online             = @($a | Where-Object { $_.online }).Count
        byOs               = @($byOs)
        versionBreakdown   = @($ver)
        newestVersion      = $newest
        onboardedThisMonth = $onboard
        offline            = $offline
        outdated           = $outdated
        agents             = $agentsList
    }
}

function Get-AlertsCount {
    param([string]$BaseUrl, [string]$TenantToken)
    $all = @(); $cursor = $null; $page = 0
    do {
        $u = if ($cursor) { "$BaseUrl/api/alert_manager/v1/alerts?limit=200&after=$([uri]::EscapeDataString($cursor))" }
             else { "$BaseUrl/api/alert_manager/v1/alerts?limit=200" }
        $raw = Invoke-Curl -Token $TenantToken -Url $u -MaxTime 90
        $j = $raw | ConvertFrom-Json
        if ($j.domain -or $j.code) { break }
        $all += $j.items; $page++
        $cursor = $j.paging.cursors.after
    } while ($cursor -and $page -lt 50)
    return @($all).Count
}

$script:ErrCodesPt = $null
function Get-ErrorDescription {
    param([string]$Code)
    if (-not $Code) { return "" }
    if ($null -eq $script:ErrCodesPt) {
        $p = Join-Path $PSScriptRoot "..\config\error_codes_pt.json"
        $script:ErrCodesPt = if (Test-Path $p) {
            [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false)) | ConvertFrom-Json
        } else { [pscustomobject]@{} }
    }
    $hit = $script:ErrCodesPt.PSObject.Properties[$Code]
    if ($hit) { return $hit.Value } else { return $Code }
}

function Measure-Alerts {
    <# Bucket security alerts by severity, category, and type for the report Security section. #>
    param([object[]]$Alerts)
    $a = @($Alerts)
    $sev = { param($s) @($a | Where-Object { $_.severity -eq $s }).Count }
    $byCat = @{}
    $a | Group-Object category | ForEach-Object { $byCat[$_.Name] = $_.Count }
    $byType = $a | Group-Object type | Sort-Object Count -Descending | ForEach-Object {
        $g = $_.Group
        $machines = @($g | ForEach-Object {
            if ($_.details -and $_.details.resourceNames) { "$($_.details.resourceNames)" }
        } | Where-Object { $_ } | Select-Object -Unique)
        $dates = @($g | ForEach-Object { $_.createdAt } | Where-Object { $_ } | Sort-Object)
        [pscustomobject]@{
            type=$_.Name; count=$_.Count; machines=$machines
            firstSeen=$(if($dates){$dates[0]}else{$null}); lastSeen=$(if($dates){$dates[-1]}else{$null})
        }
    }
    [pscustomobject]@{
        total = $a.Count
        bySeverity = [pscustomobject]@{
            critical=(& $sev 'critical'); warning=(& $sev 'warning')
            error=(& $sev 'error'); info=(& $sev 'info')
        }
        byCategory = [pscustomobject]$byCat
        byType = @($byType)
        topType = if ($byType) { $byType[0].type } else { $null }
    }
}

function Get-ServiceMap {
    param([object[]]$UsageItems)
    $on = $UsageItems | Where-Object { $_.offering_item.status -eq 1 }
    $names = @($on | ForEach-Object { "$($_.usage_name)" })
    $any = { param($rx) [bool]($names | Where-Object { $_ -match $rx }) }
    [pscustomobject]@{
        backup = & $any '^(servers|vms|workstations)$'
        edr    = & $any 'adv_security|edr'
        mdr    = & $any 'adv_security_mdr'
        va     = & $any 'adv_management|vulnerability'
        m365   = & $any 'mailbox|m365|o365|sharepoint|gsuite|gworkspace|hosted_exchange'
        dr     = & $any 'adv_dr|dr_'
    }
}

function Get-ActiveServiceMap {
    <# What a tenant is REALLY using/running: service-specific usage value>0 OR real activity.
       Device counts (workstations/servers/vms) never imply backup. #>
    param([object[]]$UsageItems, $Backup, $Security, $Vuln, $Manual)
    $on = @($UsageItems | Where-Object { $_.offering_item.status -eq 1 })
    $sumv = { param($rx) [double](($on | Where-Object { "$($_.usage_name)" -match $rx } | ForEach-Object { [double]$_.value } | Measure-Object -Sum).Sum) }
    [pscustomobject]@{
        backup = ((& $sumv '^storage$|local_storage|pack_adv_backup') -gt 0) -or ([bool]$Backup   -and [int]$Backup.runs        -gt 0)
        edr    = ((& $sumv 'pack_adv_security|adv_security')           -gt 0) -or ([bool]$Security -and [int]$Security.total      -gt 0)
        va     = ((& $sumv 'pack_adv_management|^adv_management')      -gt 0) -or ([bool]$Vuln     -and [int]$Vuln.scansRun       -gt 0)
        m365   = ((& $sumv 'mailbox|c2c_storage|o365|sharepoint|gsuite') -gt 0) -or ([bool]$Manual -and [int]$Manual.m365SeatsUsed -gt 0)
        mdr    = $false
        dr     = $false
    }
}

function Get-M365FromUsage {
    <# Extract M365 figures from a (Workspace) tenant's /usages items, for merged BIOCAL-style reports.
       Returns seats used/shared, SharePoint sites, and backup storage TB (from c2c_storage bytes). #>
    param([object[]]$UsageItems)
    $on = @($UsageItems | Where-Object { $_.offering_item.status -eq 1 })
    $sum = { param($n) [double](($on | Where-Object { $_.usage_name -eq $n } | ForEach-Object { [double]$_.value } | Measure-Object -Sum).Sum) }
    $c2c = & $sum 'c2c_storage'
    [pscustomobject]@{
        m365SeatsUsed       = [int](& $sum 'mailboxes')
        m365SeatsShared     = [int](& $sum 'm365_seats_shared')
        m365SharepointSites = [int](& $sum 'o365_sharepoint_sites')
        m365BackupStorageTB = if ($c2c -gt 0) { [math]::Round($c2c / 1TB, 2) } else { $null }
    }
}

function Get-ContractServiceMap {
    <# Derive the {backup,edr,mdr,va,m365,dr} service map from a client's contracted service lines
       (free-text 'service' strings) so the report is gated on what the client actually bought. #>
    param([object[]]$Contract)
    $svcText = @($Contract | ForEach-Object { "$($_.service)".ToUpper() })
    $any = { param($rx) [bool]($svcText | Where-Object { $_ -match $rx }) }
    # Product-name keywords -> report sections. "STANDARD PROTECTION" is a full bundle (backup+security);
    # any "...SECURITY..." line (EDR / Standard Security / Email Security) maps to the security section.
    [pscustomobject]@{
        backup = & $any 'BACKUP|STANDARD PROTECTION'
        edr    = & $any '\bEDR\b|SECURITY|STANDARD PROTECTION'
        mdr    = & $any '\bMDR\b'
        va     = & $any 'VULNERABILIT|ADVANCED MANAGEMENT'
        m365   = & $any 'M365|MICROSOFT 365'
        dr     = & $any 'DISASTER|DRAAS'
    }
}

function Get-OutOfContract {
    <# Services that are ACTIVE but NOT contracted -> "Pontos de atenção — fora do contrato".
       Only meaningful when a contract is on file. Returns [] otherwise. #>
    param([pscustomobject]$ContractMap, [pscustomobject]$ActiveMap, $Backup, $Security, $Vuln, $Manual)
    if (-not $ContractMap) { return @() }
    $scope = @()
    if ($ContractMap.backup) { $scope += 'Backup' }
    if ($ContractMap.edr)    { $scope += 'EDR' }
    if ($ContractMap.va)     { $scope += 'Gerenciamento/VA' }
    if ($ContractMap.m365)   { $scope += 'M365' }
    $scopeStr = if (@($scope).Count) { $scope -join ', ' } else { 'nenhum serviço contratado' }
    $msg = { param($lbl) "$lbl ativo neste tenant, porém fora do contrato ($scopeStr). Revisar cobrança/escopo." }
    $out = @()
    if ($ActiveMap.backup -and -not $ContractMap.backup) {
        $topErr = @(@($Backup.errorReasons) | Select-Object -First 3 | ForEach-Object { [pscustomobject]@{ labelPt=$_.plainPt; count=$_.count } })
        $out += [pscustomobject]@{ service='backup'; labelPt='Backup';
            metrics=[pscustomobject]@{ runs=$Backup.runs; ok=$Backup.ok; error=$Backup.error; successPct=$Backup.successPct; topErrors=$topErr };
            message=(& $msg 'Backup') }
    }
    if ($ActiveMap.edr -and -not $ContractMap.edr) {
        $topT = @(@($Security.byType) | Select-Object -First 3 | ForEach-Object { [pscustomobject]@{ labelPt=$_.type; count=$_.count } })
        $out += [pscustomobject]@{ service='edr'; labelPt='EDR';
            metrics=[pscustomobject]@{ total=$Security.total; topTypes=$topT };
            message=(& $msg 'EDR') }
    }
    if ($ActiveMap.va -and -not $ContractMap.va) {
        $out += [pscustomobject]@{ service='va'; labelPt='Avaliação de vulnerabilidades';
            metrics=[pscustomobject]@{ scansRun=$Vuln.scansRun };
            message=(& $msg 'Avaliação de vulnerabilidades') }
    }
    if ($ActiveMap.m365 -and -not $ContractMap.m365) {
        $out += [pscustomobject]@{ service='m365'; labelPt='Microsoft 365';
            metrics=[pscustomobject]@{ seatsUsed=$Manual.m365SeatsUsed; storageTB=$Manual.m365BackupStorageTB };
            message=(& $msg 'Microsoft 365') }
    }
    return @($out)
}

function ConvertFrom-AuditResponse {
    <# Parse raw audit API response string; sets accessPending=$true on AccessDeniedError or parse failure. #>
    param([string]$Raw)
    $j = $null
    try { $j = $Raw | ConvertFrom-Json } catch { return [pscustomobject]@{ accessPending=$true; events=@() } }
    if ($j.code -eq "AccessDeniedError" -or $j.domain -eq "Access") {
        return [pscustomobject]@{ accessPending=$true; events=@() }
    }
    return [pscustomobject]@{ accessPending=$false; events=@($j.items) }
}

function Get-AuditEvents {
    <# POST to audit search endpoint and return parsed ConvertFrom-AuditResponse result. #>
    param([string]$BaseUrl,[string]$Token,[string]$TenantId,[string]$StartIso,[string]$EndIso)
    $body = (@{ tenant_id=$TenantId; from=$StartIso; to=$EndIso } | ConvertTo-Json -Compress)
    $url = "$BaseUrl/api/audit/v2/events/search"
    $raw = & $script:CurlBin -s --max-time 90 -X POST -H "Authorization: Bearer $Token" -H "Content-Type: application/json" --data $body $url
    return ConvertFrom-AuditResponse $raw
}

function Get-VaActivities {
    param([string]$BaseUrl,[string]$TenantToken,[string]$StartIso,[string]$EndIso)
    $u = "$BaseUrl/api/task_manager/v2/activities?policyType=vulnerability_assessment&completedAt=gt($StartIso)&completedAt=lt($EndIso)&limit=200"
    $raw = Invoke-Curl -Token $TenantToken -Url $u -MaxTime 280
    $j = $raw | ConvertFrom-Json
    return @($j.items)
}

function Measure-Va {
    param([object[]]$VaActivities)
    $a = @($VaActivities)
    $last = ($a | Where-Object { $_.completedAt } | Sort-Object completedAt -Descending | Select-Object -First 1)
    [pscustomobject]@{
        scansRun     = $a.Count
        scanFailures = @($a | Where-Object { $_.result.code -eq "error" }).Count
        lastScanDate = if ($last) { $last.completedAt } else { $null }
    }
}

function Get-VaFindings {
    param([string]$BaseUrl,[string]$TenantToken)
    # software_management/v4 is a live service; exact findings resource pending confirmation.
    # Until confirmed, return a graceful placeholder (consumed by the report's fallback note).
    return [pscustomobject]@{ available=$false; findings=@(); note="Endpoint de findings de VA pendente de confirmacao." }
}

$script:AlertTypesPt = $null
function Get-AlertBuckets {
    param([pscustomobject]$Security)
    if ($null -eq $script:AlertTypesPt) {
        $p = Join-Path $PSScriptRoot "..\config\alert_types_pt.json"
        $script:AlertTypesPt = if (Test-Path $p) { [System.IO.File]::ReadAllText($p,[System.Text.UTF8Encoding]::new($false)) | ConvertFrom-Json } else { [pscustomobject]@{} }
    }
    $posture=@(); $hygiene=@()
    foreach ($t in @($Security.byType)) {
        $m = $script:AlertTypesPt.PSObject.Properties["$($t.type)"]
        if ($m) { $bucket=$m.Value.bucket; $label=$m.Value.labelPt }
        else {
            $bucket = if ("$($t.type)" -match 'Offline|Malicious|Edr|Detection|Device|Blocked|Ransom') { 'posture' } else { 'hygiene' }
            $label = "$($t.type)"
        }
        $entry = [pscustomobject]@{ labelPt=$label; type=$t.type; count=$t.count; machines=@($t.machines) }
        if ($bucket -eq 'posture') { $posture += $entry } else { $hygiene += $entry }
    }
    [pscustomobject]@{ posture=@($posture); hygiene=@($hygiene) }
}

function Get-OfferingCategory {
    <# Classify an offering by usage_name into a report category. Device counts are shared
       (always shown); only backup/edr/va/m365 are gated by the service map. #>
    param([string]$UsageName)
    $u = "$UsageName"
    if ($u -match '^(workstations|servers|vms)$')                              { return 'device' }
    if ($u -match 'mailbox|^m365|o365|c2c_storage|sharepoint|gsuite|gworkspace|hosted_exchange|octiga') { return 'm365' }
    if ($u -match 'pack_adv_management|^adv_management|vulnerabilit')          { return 'va' }
    if ($u -match 'pack_adv_security|adv_security|edr')                        { return 'edr' }
    if ($u -match '^storage$|local_storage|archiving_storage|pack_adv_backup') { return 'backup' }
    return 'other'
}

$script:UsageLabelsPt = $null
function Get-ServiceUsage {
    param([object[]]$UsageItems)
    if ($null -eq $script:UsageLabelsPt) {
        $p = Join-Path $PSScriptRoot "..\config\usage_labels_pt.json"
        $script:UsageLabelsPt = if (Test-Path $p) { [System.IO.File]::ReadAllText($p,[System.Text.UTF8Encoding]::new($false)) | ConvertFrom-Json } else { [pscustomobject]@{} }
    }
    $out = @()
    foreach ($it in @($UsageItems | Where-Object { $_.offering_item.status -eq 1 })) {
        $name  = "$($it.name)"
        $uname = "$($it.usage_name)"
        $quota = $it.offering_item.quota.value
        $used  = [double]$it.value
        $isBytes = ("$($it.measurement_unit)" -eq 'bytes')
        # Convert byte-valued offerings (cloud/local storage) to BINARY GB to match the console.
        if ($isBytes) {
            $used = [math]::Round($used / 1GB)
            if ($quota -and [double]$quota -gt 0) { $quota = [math]::Round([double]$quota / 1GB) }
        }
        # Drop pure-noise rows: nothing used AND nothing contracted (e.g. c2c_storage 0/null).
        if ($used -eq 0 -and (-not $quota -or [double]$quota -le 0)) { continue }
        $lab   = $script:UsageLabelsPt.PSObject.Properties[$uname]
        $label = if ($lab) { $lab.Value } else { (Get-Culture).TextInfo.ToTitleCase(($uname -replace '_',' ')) }
        if ($isBytes -and $label -notmatch 'GB') { $label = "$label (GB)" }
        $pct   = $null; $status = 'ok'
        if ($quota -and [double]$quota -gt 0) {
            $pct = [int]([math]::Round(100.0*$used/[double]$quota))
            if ($pct -ge 90) { $status = 'warn' }
        }
        $out += [pscustomobject]@{
            name=$name; labelPt=$label; quota=$quota; used=$used; pct=$pct; status=$status
            category=(Get-OfferingCategory $uname)
        }
    }
    return @($out)
}

Export-ModuleMember -Function Get-AcronisToken, Get-AcronisCustomers, Get-AcronisUsages,
    Get-ServiceProfile, Get-MonthBounds, Get-BackupActivities, Measure-Backup,
    Measure-Agents, Get-AlertsCount, Get-ErrorDescription, Get-ServiceMap, Measure-Alerts,
    ConvertFrom-AuditResponse, Get-AuditEvents,
    Get-VaActivities, Measure-Va, Get-VaFindings, Invoke-Curl, Get-PagedItems,
    Get-ServiceUsage, Get-AlertBuckets, Get-ContractServiceMap, Get-M365FromUsage,
    Get-OfferingCategory, Get-ActiveServiceMap, Get-OutOfContract
