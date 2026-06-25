param([Parameter(Mandatory)][string]$FactsPath,[Parameter(Mandatory)][string]$OutDir,[Parameter(Mandatory)][string]$TemplateDir)
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "lib/AcronisReport.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "lib/AcronisRender.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "lib/AcronisHtml.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "lib/AcronisDeck.psm1") -Force -DisableNameChecking
# Re-import AcronisReport after AcronisDeck so its exports remain visible in this scope
# (AcronisDeck imports AcronisReport internally as a nested module, which shadows the global export —
# same pattern as Run-MonthlyReports.ps1 uses for AcronisAnalysis after AcronisFacts).
Import-Module (Join-Path $PSScriptRoot "lib/AcronisReport.psm1") -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot "lib/AcronisDeckPdf.psm1") -Force -DisableNameChecking
$cfg   = Join-Path $PSScriptRoot "config"
$sla   = if (Test-Path (Join-Path $cfg "sla.json")) { (Read-Utf8 (Join-Path $cfg "sla.json")) | ConvertFrom-Json } else { [pscustomobject]@{backupSuccessPct=95;agentsOnlinePct=98;storageOccupancyMaxPct=85} }
$facts = (Read-Utf8 $FactsPath) | ConvertFrom-Json
$slug  = ($facts.meta.client -replace '[^\w\- ]','' -replace '\s+','_').ToUpper()
$base  = "RELATORIO_MSSP_${slug}_$($facts.meta.month)"
# history
$histDir = Join-Path $PSScriptRoot "output/history"; New-Item -ItemType Directory -Force $histDir | Out-Null
$histFile = Join-Path $histDir "$($facts.meta.tenantId).json"
$history = if (Test-Path $histFile) { @((Read-Utf8 $histFile) | ConvertFrom-Json) } else { @() }
# render
$html  = Build-HtmlReport $facts $history $sla $cfg
$md    = Build-Report $facts          # unchanged markdown artifact
$gamma = Build-GammaPrompt $facts
Write-Utf8 (Join-Path $OutDir "$base.md") $md
Write-Utf8 (Join-Path $OutDir "$base.html") $html
Write-Utf8 (Join-Path $OutDir "GAMMA_PROMPT_${slug}_$($facts.meta.month).txt") $gamma
$pdf = Convert-HtmlToPdf (Join-Path $OutDir "$base.html") (Join-Path $OutDir "$base.pdf")
# Gamma-style slide deck (HTML) + landscape PDF — gated to contracted services
$deckHtmlPath = Join-Path $OutDir "$base.deck.html"
Write-Utf8 $deckHtmlPath (Build-DeckHtml $facts $cfg)
$deckPdf = Convert-DeckToPdf $deckHtmlPath (Join-Path $OutDir "$base.deck.pdf")
# append history (after successful render)
$entry = [pscustomobject]@{ month=$facts.meta.month; successPct=$(if($facts.backup){$facts.backup.successPct}else{$null}); cloudGB=$(if($facts.storage){$facts.storage.cloudGB}else{$null}); occupancyPct=$(if($facts.storage){$facts.storage.occupancyPct}else{$null}); protectedWorkloads=$facts.profile.protectedWorkloads; alertsCritical=$(if($facts.security){$facts.security.bySeverity.critical}else{$null}); healthScore=$facts.analysis.healthScore }
$history = @($history | Where-Object { $_.month -ne $facts.meta.month }) + $entry
($history | ConvertTo-Json -Depth 5) | Out-File -Encoding utf8 $histFile
Write-Host "  report: md+html$(if($pdf.pdf){"+pdf ($($pdf.tool))"}else{''}) + deck.html$(if($deckPdf){'+deck.pdf'}else{''}) + gamma"
