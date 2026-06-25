# automation/lib/AcronisAnalysis.psm1
$ErrorActionPreference = "Stop"
function Get-CriticalIncidents {
    param([pscustomobject]$Backup, [int]$MinFailures = 5)
    $topCause = ($Backup.errorReasons | Sort-Object count -Descending | Select-Object -First 1)
    $causeTxt = if ($topCause) { $topCause.plainPt } else { "causa não identificada" }
    $out = @()
    foreach ($m in $Backup.perMachine) {
        if ($m.err -gt 0 -and $m.err -eq $m.runs) {
            $out += [pscustomobject]@{ machine=$m.machine; failures=$m.err; pattern="100% failure"; cause=$causeTxt }
        } elseif ($m.err -ge $MinFailures) {
            $out += [pscustomobject]@{ machine=$m.machine; failures=$m.err; pattern="high failure count"; cause=$causeTxt }
        }
    }
    return @($out | Sort-Object failures -Descending)
}
function Get-Analysis {
    param([pscustomobject]$Facts)
    $pri = @()
    foreach ($i in @($Facts.incidents)) {
        $pri += "CRÍTICO: $($i.machine) - $($i.pattern) ($($i.cause))"
    }
    if ($Facts.storage -and $Facts.storage.occupancyPct -ge 80) {
        $pri += "Capacidade: armazenamento em $($Facts.storage.occupancyPct)% do contratado"
    }
    if ($Facts.lifecycle -and @($Facts.lifecycle.outdated).Count -gt 0) {
        $pri += "Atualizar $(@($Facts.lifecycle.outdated).Count) agentes desatualizados"
    }
    if ($Facts.lifecycle -and @($Facts.lifecycle.offline).Count -gt 0) {
        $pri += "Investigar $(@($Facts.lifecycle.offline).Count) agentes offline"
    }
    # health score: start 100, subtract for failures, incidents, capacity, criticals
    $score = 100
    if ($Facts.backup) { $score -= [int]((100 - [double]$Facts.backup.successPct) * 0.5) }
    $score -= 10 * @($Facts.incidents).Count
    if ($Facts.storage -and $Facts.storage.occupancyPct -ge 80) { $score -= 5 }
    if ($Facts.security) { $score -= 3 * [int]$Facts.security.bySeverity.critical }
    if ($score -lt 0) { $score = 0 }
    if ($score -gt 100) { $score = 100 }
    [pscustomobject]@{ healthScore=[int]$score; topPriorities=@($pri) }
}
Export-ModuleMember -Function Get-CriticalIncidents, Get-Analysis