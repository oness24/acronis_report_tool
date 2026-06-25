<#
.SYNOPSIS
  Monthly MSSP report pipeline for Contego Security (Acronis Cyber Protect Cloud).
  READ-ONLY. Auto-discovers every customer on every configured data center, collects
  usage + backup activity + agents + alerts, computes aggregates, and drafts a PT-BR
  report + Gamma prompt per client. Cross-platform (Windows PowerShell 5.1 / pwsh 7 on Linux/Colab).

.EXAMPLE
  ./Run-MonthlyReports.ps1                                # defaults to LAST month (run in June -> reports May)
  ./Run-MonthlyReports.ps1 -Verify                        # last month, activities collected twice and diffed
  ./Run-MonthlyReports.ps1 -Month 2026-03                 # override for a back-fill
  ./Run-MonthlyReports.ps1 -OnlyTenant 3c0c9333-ee22-47a1-9fbc-46f0a9e8f434
#>
param(
    # Defaults to the PREVIOUS calendar month (standard MSSP cadence: run in June, report May).
    [ValidatePattern('^\d{4}-\d{2}$')][string]$Month = (Get-Date).AddMonths(-1).ToString('yyyy-MM'),
    [switch]$Verify,
    [string]$OnlyTenant,
    # Discover customers on every data center and emit a JSON list (id/name/dataCenter)
    # between markers, then exit BEFORE any heavy collection. Used by the Colab
    # notebook to drive a resilient one-client-at-a-time loop.
    [switch]$ListTenants,
    [string]$Root = "$PSScriptRoot"
)
$ErrorActionPreference = "Stop"
Import-Module "$Root/lib/AcronisReport.psm1" -Force
Import-Module "$Root/lib/AcronisAnalysis.psm1" -Force
Import-Module "$Root/lib/AcronisFacts.psm1" -Force
# Re-import AcronisAnalysis after AcronisFacts so its exports remain visible in caller scope
# (AcronisFacts imports AcronisAnalysis internally as a nested module, which shadows the global export)
Import-Module "$Root/lib/AcronisAnalysis.psm1" -Force
Import-Module "$Root/lib/AcronisPptx.psm1" -Force

$dcs       = (Get-Content "$Root/config/datacenters.json" -Raw | ConvertFrom-Json).datacenters
$secrets   = Get-Content "$Root/config/secrets.json" -Raw | ConvertFrom-Json
$contracts = Get-Content "$Root/config/contracts.json" -Raw | ConvertFrom-Json
$bounds    = Get-MonthBounds -Month $Month
Write-Host "Report month: $Month  (window $($bounds.StartIso) -> $($bounds.EndIso), UTC)" -ForegroundColor Cyan

$outData = "$Root/output/$Month/data";    New-Item -ItemType Directory -Force $outData | Out-Null
$outRep  = "$Root/output/$Month/reports"; New-Item -ItemType Directory -Force $outRep  | Out-Null

# Resolve Python launcher once (optional -- absent means no .pptx deck, not a failure).
# Skip the probe entirely on -ListTenants (enumeration only; no generation -> no warning).
$pyLauncher = if (-not $ListTenants) { Resolve-PythonLauncher } else { $null }
if (-not $ListTenants) {
    if ($pyLauncher) { Write-Host "  pptx: Python launcher encontrado: $($pyLauncher -join ' ')" -ForegroundColor DarkCyan }
    else             { Write-Warning "  pptx: Python/python-pptx nao encontrado - decks PPTX serao pulados" }
}

# Tenants that are merged INTO a primary client (e.g. "Biocal Workspace" -> BIOCAL) get no standalone report.
$mergeTargetIds = @()
$contracts.PSObject.Properties | Where-Object { $_.Name -ne '_comment' } | ForEach-Object {
    foreach ($mf in @($_.Value.mergeFrom)) { if ($mf.tenantId) { $mergeTargetIds += $mf.tenantId } }
}

function Slug([string]$s) { ($s.Trim() -replace '[^\w\- ]','' -replace '\s+','_').ToUpper() }

$summary    = @()
$discovered = @()

foreach ($dc in $dcs) {
    $cred = $secrets.($dc.name)
    if (-not $cred -or $cred.secret -like 'PUT-*') { Write-Warning "[$($dc.name)] no creds in secrets.json - skipping"; continue }
    Write-Host "`n########## DATA CENTER: $($dc.name) ##########" -ForegroundColor Cyan

    $partnerToken = Get-AcronisToken -BaseUrl $dc.baseUrl -ClientId $cred.clientId -Secret $cred.secret
    $customers = Get-AcronisCustomers -BaseUrl $dc.baseUrl -PartnerToken $partnerToken -RootTenantId $dc.rootTenantId
    if ($OnlyTenant) { $customers = $customers | Where-Object { $_.id -eq $OnlyTenant } }
    $customers = @($customers)
    Write-Host "Discovered $($customers.Count) customer tenant(s)."

    # -ListTenants: just record reportable tenants (skip merge targets) and move on.
    if ($ListTenants) {
        foreach ($c in $customers) {
            if ($mergeTargetIds -contains $c.id) { continue }
            $discovered += [pscustomobject]@{ id = $c.id; name = $c.name; dataCenter = $dc.name }
        }
        continue
    }

    foreach ($c in $customers) {
        if ($mergeTargetIds -contains $c.id) { Write-Host "`n--- skip $($c.name) (merged into a primary client) ---" -ForegroundColor DarkGray; continue }
        Write-Host "`n=== $($c.name) [$($c.id)] ===" -ForegroundColor Yellow
        try {
            $usages   = Get-AcronisUsages -BaseUrl $dc.baseUrl -PartnerToken $partnerToken -TenantId $c.id
            $svcUsage = Get-ServiceUsage $usages
            # Service maps:
            #  - enabledMap : offerings present (status=1) -> COLLECTION scope.
            #  - contractMap: what the client bought (strict section gating when on file).
            #  - activeMap  : computed AFTER collection (real usage + activity) -> sections for
            #                 no-contract clients, and the out-of-contract comparison.
            $enabledMap = Get-ServiceMap $usages
            $contractLines = $contracts.($c.id).contract
            $contractMap = $null
            if ($contractLines -and @($contractLines).Count) {
                $contractMap = Get-ContractServiceMap $contractLines; $svcSource = 'contract'
                Write-Host "  gating: contrato ($(@($contractLines).Count) linhas), estrito + fora do contrato" -ForegroundColor DarkCyan
            } else { $svcSource = 'usages' }
            # Collection scope = contracted OR platform-enabled (so out-of-contract activity is captured).
            $collect = [pscustomobject]@{
                backup = ($contractMap -and $contractMap.backup) -or $enabledMap.backup
                edr    = ($contractMap -and $contractMap.edr)    -or $enabledMap.edr
                va     = ($contractMap -and $contractMap.va)     -or $enabledMap.va
                m365   = ($contractMap -and $contractMap.m365)   -or $enabledMap.m365
            }
            $tToken   = Get-AcronisToken -BaseUrl $dc.baseUrl -ClientId $cred.clientId -Secret $cred.secret -TenantId $c.id

            $backup = $null; [object[]]$incidents = @()
            if ($collect.backup) {
                Write-Host "  backup -> collecting activities..."
                $acts = Get-BackupActivities -BaseUrl $dc.baseUrl -TenantToken $tToken -StartIso $bounds.StartIso -EndIso $bounds.EndIso
                $backup = Measure-Backup $acts
                if ($Verify) {
                    $acts2 = Get-BackupActivities -BaseUrl $dc.baseUrl -TenantToken $tToken -StartIso $bounds.StartIso -EndIso $bounds.EndIso
                    $b2 = Measure-Backup $acts2
                    $backup | Add-Member verifyMatches ([bool]($b2.runs -eq $backup.runs -and $b2.ok -eq $backup.ok -and $b2.error -eq $backup.error))
                }
                [object[]]$incidents = @(Get-CriticalIncidents $backup)
            } else { Write-Host "  no backup offering (security-only) -> skipping activities" }

            $security = $null
            if ($collect.edr) {
                Write-Host "  edr -> collecting alerts..."
                $alerts = Get-PagedItems -Token $tToken -Url "$($dc.baseUrl)/api/alert_manager/v1/alerts?limit=200" -MaxTime 120
                $security = Measure-Alerts $alerts
            }

            $vuln = $null
            if ($collect.va) {
                Write-Host "  va -> collecting VA activities..."
                $vuln = Measure-Va (Get-VaActivities -BaseUrl $dc.baseUrl -TenantToken $tToken -StartIso $bounds.StartIso -EndIso $bounds.EndIso)
            }

            Write-Host "  collecting agents..."
            $lifecycle = Measure-Agents (Get-PagedItems -Token $tToken -Url "$($dc.baseUrl)/api/agent_manager/v2/agents" -MaxTime 120) $Month

            Write-Host "  collecting audit events..."
            $audit = Get-AuditEvents -BaseUrl $dc.baseUrl -Token $tToken -TenantId $c.id -StartIso $bounds.StartIso -EndIso $bounds.EndIso

            # storage (binary GB) + occupancy vs contracted cloud-storage line, if any
            $sumBytes = { param($n) ($usages | Where-Object { $_.usage_name -eq $n } | ForEach-Object { [double]$_.value } | Measure-Object -Sum).Sum }
            $cloudGB = [math]::Round([double](& $sumBytes "storage")/1GB,2)
            $contractCloudGB = (($contracts.($c.id).contract | Where-Object { $_.service -match 'CLOUD STORAGE' } | Measure-Object qty -Sum).Sum)
            $storage = [pscustomobject]@{
                cloudGB      = $cloudGB
                localGB      = [math]::Round([double](& $sumBytes "local_storage_total")/1GB,2)
                contractedGB = $contractCloudGB
                occupancyPct = if ($contractCloudGB) { [math]::Round(100.0*$cloudGB/$contractCloudGB,1) } else { 0 }
            }

            # Cross-DC merge: pull each secondary tenant (e.g. "Biocal Workspace" M365) and fold it in.
            $manual = $contracts.($c.id).manual
            foreach ($mf in @($contracts.($c.id).mergeFrom)) {
                if (-not $mf -or -not $mf.tenantId) { continue }   # skip the PS5.1 @($null) phantom element
                $mDc = $dcs | Where-Object { $_.name -eq $mf.dataCenter } | Select-Object -First 1
                $mCred = $secrets.($mf.dataCenter)
                if (-not $mDc -or -not $mCred -or $mCred.secret -like 'PUT-*') { Write-Warning "  mergeFrom $($mf.dataCenter) unreachable - skipping"; continue }
                Write-Host "  merging tenant $($mf.label) [$($mf.dataCenter)] ..." -ForegroundColor DarkCyan
                $mTok = Get-AcronisToken -BaseUrl $mDc.baseUrl -ClientId $mCred.clientId -Secret $mCred.secret
                $mUsages = Get-AcronisUsages -BaseUrl $mDc.baseUrl -PartnerToken $mTok -TenantId $mf.tenantId
                $svcUsage = @($svcUsage) + @(Get-ServiceUsage $mUsages)   # secondary offerings join the service table
                $m365 = Get-M365FromUsage $mUsages
                if (-not $manual) { $manual = [pscustomobject]@{} }
                $manual | Add-Member -NotePropertyName m365SeatsUsed       -NotePropertyValue $m365.m365SeatsUsed       -Force
                $manual | Add-Member -NotePropertyName m365SeatsShared     -NotePropertyValue $m365.m365SeatsShared     -Force
                $manual | Add-Member -NotePropertyName m365SharepointSites -NotePropertyValue $m365.m365SharepointSites -Force
                $manual | Add-Member -NotePropertyName m365BackupStorageTB -NotePropertyValue $m365.m365BackupStorageTB -Force
            }

            # Real activity map (post-collection) + final section map + out-of-contract findings.
            $activeMap = Get-ActiveServiceMap $usages $backup $security $vuln $manual
            if ($contractMap) {
                $svc = $contractMap
                [object[]]$outOfContract = @(Get-OutOfContract $contractMap $activeMap $backup $security $vuln $manual)
            } else {
                $svc = $activeMap
                [object[]]$outOfContract = @()
            }

            $facts = New-ClientFacts @{
                meta = [pscustomobject]@{ client=$c.name.Trim(); tenantId=$c.id; dataCenter=$dc.name; month=$Month;
                        windowUtc="$($bounds.StartIso)..$($bounds.EndIso)"; generatedAt=(Get-Date).ToUniversalTime().ToString("u"); serviceMapSource=$svcSource }
                serviceMap=$svc; serviceUsage=$svcUsage; backup=$backup; incidents=$incidents; security=$security; vulnerability=$vuln;
                lifecycle=$lifecycle; audit=$audit; storage=$storage;
                contract=($contracts.($c.id).contract); manual=$manual; outOfContract=$outOfContract
            }
            $slug = Slug $c.name
            $facts | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 "$outData/facts_$slug.json"
            Write-Host "  facts written -> $outData/facts_$slug.json" -ForegroundColor Green

            try {
                & "$Root/Generate-Report.ps1" -FactsPath "$outData/facts_$slug.json" -OutDir $outRep -TemplateDir "$Root/templates"
            } catch {
                Write-Warning "  Generate-Report failed (non-fatal): $($_.Exception.Message)"
            }

            Invoke-PptxBuild -FactsPath "$outData/facts_$slug.json" -OutDir $outRep -ConfigDir "$Root/config" -Launcher $pyLauncher

            $summary += [pscustomobject]@{
                Client=$c.name; DC=$dc.name; Backup=$(if($backup){"$($backup.successPct)%"}else{"-"})
                ServiceMap="$(if($svc.backup){'BK'})$(if($svc.edr){'+EDR'})$(if($svc.va){'+VA'})"
                CloudGB=$storage.cloudGB; OccupancyPct=$storage.occupancyPct
            }
            Write-Host "  OK -> facts + report drafted" -ForegroundColor Green
        } catch {
            Write-Warning "  FAILED for $($c.name): $($_.Exception.Message)"
            $summary += [pscustomobject]@{ Client=$c.name; DC=$dc.name; Backup="ERROR"; Workloads=""; CloudGB=""; Alerts="" }
        }
    }
}

# -ListTenants: emit the discovered tenants as a JSON array between markers, then stop.
if ($ListTenants) {
    $json = '[' + (($discovered | ForEach-Object { $_ | ConvertTo-Json -Depth 4 -Compress }) -join ',') + ']'
    Write-Host "<<<TENANTS_JSON>>>"
    Write-Output $json
    Write-Host "<<<END_TENANTS_JSON>>>"
    return
}

Write-Host "`n================ RUN SUMMARY ($Month) ================" -ForegroundColor Cyan
$summary | Format-Table -AutoSize
$summary | ConvertTo-Json -Depth 4 | Out-File -Encoding utf8 "$Root/output/$Month/run_summary.json"
Write-Host "Reports in: $outRep"
