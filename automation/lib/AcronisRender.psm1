$ErrorActionPreference = "Stop"
function Read-Utf8([string]$p) { [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false)) }
function Write-Utf8([string]$p,[string]$t) { [System.IO.File]::WriteAllText($p, $t, (New-Object System.Text.UTF8Encoding($false))) }
function New-MdTable {
    param([string[]]$Headers, [object[]]$Rows)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("| " + ($Headers -join " | ") + " |")
    [void]$sb.AppendLine("|" + (($Headers | ForEach-Object { "---" }) -join "|") + "|")
    foreach ($r in $Rows) { [void]$sb.AppendLine("| " + (@($r) -join " | ") + " |") }
    return $sb.ToString()
}
function Get-StatusWord {
    param([double]$pct)
    if ($pct -ge 90) { "OK" } elseif ($pct -ge 70) { "Atenção" } else { "Crítico" }
}
function Build-ExecSummary {
    param([pscustomobject]$Facts)
    $rows = @()
    $rows += ,@("Saúde geral", "$($Facts.analysis.healthScore)/100")
    if ($Facts.serviceMap.backup) {
        $rows += ,@("Backup - taxa de sucesso", "$($Facts.backup.successPct)% (excl. pior: $($Facts.backup.successPctExclWorst)%)")
        $rows += ,@("Armazenamento em nuvem", "$($Facts.storage.cloudGB) GB ($($Facts.storage.occupancyPct)% do contratado)")
    }
    if ($Facts.serviceMap.edr -and $Facts.security) {
        $rows += ,@("Alertas críticos em aberto", "$($Facts.security.bySeverity.critical)")
    }
    $prio = if (@($Facts.analysis.topPriorities).Count) { $Facts.analysis.topPriorities[0] } else { "Nenhuma ação crítica" }
    $rows += ,@("Ação prioritária", $prio)
    "## 1. Resumo executivo`n`n" + (New-MdTable @("Indicador","Valor") $rows)
}

function Build-Environment {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.backup) { return "" }
    $l = $Facts.lifecycle
    $os = (@($l.byOs) | ForEach-Object { "$($_.count) $($_.os)" }) -join ", "
    @"
## 2. Visão geral do ambiente
- Agentes instalados: $($l.total) ($os) - $($l.online) online.
- Frota: versão mais recente $($l.newestVersion); $(@($l.outdated).Count) desatualizado(s).
- Armazenamento em nuvem: $($Facts.storage.cloudGB) GB; local: $($Facts.storage.localGB) GB.

"@
}

function Build-Backup {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.backup) { return "" }
    $b = $Facts.backup
    $rows = @(
        ,@("Execuções (nível superior)", "$($b.runs) ($($b.scheduled) agendadas, $($b.manual) manuais)")
        ,@("Sucesso", "$($b.ok) ($($b.successPct)%)")
        ,@("Erros", "$($b.error)")
        ,@("Avisos", "$($b.warning)")
        ,@("Volume transferido", "~$($b.transferredGB) GB")
    )
    $reasons = (@($b.errorReasons) | ForEach-Object { "- $($_.plainPt) (x$($_.count))" }) -join "`n"
    "## 3. Atividade de backup`n`n" + (New-MdTable @("Indicador","Valor") $rows) +
      $(if ($reasons) { "`n**Causas de erro:**`n$reasons`n" } else { "" })
}

function Build-Incidents {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.backup -or -not @($Facts.incidents).Count) { return "" }
    # GROUP by pattern+cause so an all-fail month is one line per cause, not one per machine
    $groups = @($Facts.incidents) | Group-Object { "$($_.pattern)|$($_.cause)" } | Sort-Object Count -Descending
    $lines = foreach ($g in $groups) {
        $first = $g.Group[0]
        $machines = @($g.Group | ForEach-Object { $_.machine })
        $sample = ($machines | Select-Object -First 5) -join ", "
        $more = if ($machines.Count -gt 5) { " (+$($machines.Count - 5) outras)" } else { "" }
        "- **$($g.Count) máquina(s)** - $($first.pattern). Causa: $($first.cause). Ex.: $sample$more"
    }
    "## 4. Pontos críticos`n`n" + ($lines -join "`n") + "`n"
}

function Build-Security {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.edr -or -not $Facts.security) { return "" }
    $s = $Facts.security
    $rows = @(
        ,@("Total de alertas", "$($s.total)")
        ,@("Críticos", "$($s.bySeverity.critical)")
        ,@("Avisos", "$($s.bySeverity.warning)")
        ,@("Tipo mais frequente", "$($s.topType)")
    )
    "## 5. Segurança e EDR`n`n" + (New-MdTable @("Indicador","Valor") $rows)
}

function Build-Va {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.va) { return "" }
    $v = $Facts.vulnerability
    $body = "- Varreduras no mês: $($v.scansRun) ($($v.scanFailures) com falha). Última: $($v.lastScanDate)."
    "## 6. Avaliação de vulnerabilidades`n`n$body`n"
}

function Build-M365 {
    param([pscustomobject]$Facts)
    if (-not $Facts.serviceMap.m365) { return "" }
    $m = $Facts.manual
    if ($m -and $m.m365SeatsUsed) {
        $rows = @(
            ,@("Seats M365 protegidos", "$($m.m365SeatsUsed) de $($m.m365SeatsContracted)")
            ,@("Storage de backup M365", "$($m.m365BackupStorageTB) TB")
        )
        "## 7. Proteção Microsoft 365`n`n" + (New-MdTable @("Indicador","Valor") $rows)
    } else {
        "## 7. Proteção Microsoft 365`n`n> Dados do console M365 pendentes (tenant dedicado). Preencher em config/contracts.json.`n"
    }
}

function Build-Audit {
    param([pscustomobject]$Facts)
    if (-not $Facts.audit -or $Facts.audit.accessPending) {
        return "## 8. Alterações de licença e configuração`n`n> Acesso pendente: Audit log não habilitado para o cliente da API. Conceda o papel de auditoria para popular esta seção.`n"
    }
    $rows = @($Facts.audit.events) | ForEach-Object {
        ,@($_.date, $_.initiator, $_.action, $_.object, $_.method)
    }
    "## 8. Alterações de licença e configuração`n`n" + (New-MdTable @("Data","Autor","Ação","Objeto","Método") $rows)
}

function Build-Contract {
    param([pscustomobject]$Facts)
    if (-not $Facts.contract -or -not @($Facts.contract).Count) { return "" }
    $rows = @($Facts.contract) | ForEach-Object {
        $used = if ($_.PSObject.Properties['used']) { $_.used } else { "-" }
        $occ  = if ($_.PSObject.Properties['occupancyPct']) { "$($_.occupancyPct)%" } else { "-" }
        ,@($_.service, $_.qty, $used, $occ)
    }
    "## 9. Contratado vs. utilizado`n`n" + (New-MdTable @("Serviço","Contratado","Utilizado","Ocupação") $rows)
}

function Build-ActionPlan {
    param([pscustomobject]$Facts, [int]$Max = 10)
    # de-duplicate identical priority strings, then cap (an 87-line action plan is unusable)
    $items = @($Facts.analysis.topPriorities) | Select-Object -Unique
    $shown = @($items | Select-Object -First $Max)
    if ($shown.Count) {
        $n = 0
        $body = ($shown | ForEach-Object { $n++; "$n. $_" }) -join "`n"
        if ($items.Count -gt $Max) { $body += "`n... (+$($items.Count - $Max) itens adicionais)" }
    } else { $body = "Nenhuma ação prioritária identificada no período." }
    "## 10. Plano de ação`n`n$body`n"
}

function Build-Methodology {
    param([pscustomobject]$Facts)
    @"
---
*Método: execuções consideram apenas atividades de nível superior (IsProcessRoot) do tipo backup, janela $($Facts.meta.windowUtc) (UTC). Volume = bytesSaved pós dedup/compressão. Armazenamento em base binária (1 GB = 1024³ bytes), idêntico ao console Acronis. Coleta somente-leitura. Retenção do audit log: 180 dias.*
"@
}

function Convert-MarkdownToHtml {
    param([string]$Md)
    $lines = $Md -split "`r?`n"
    $html = [System.Text.StringBuilder]::new()
    $inTable = $false; $inList = $false
    function Inline([string]$s) {
        $s = $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
        $s = [regex]::Replace($s, '\*\*(.+?)\*\*', '<strong>$1</strong>')
        $s = [regex]::Replace($s, '`(.+?)`', '<code>$1</code>')
        return $s
    }
    foreach ($ln in $lines) {
        if ($ln -match '^\|(.+)\|\s*$') {
            $cells = ($ln.Trim().Trim('|') -split '\|') | ForEach-Object { $_.Trim() }
            if ($cells -and ($cells -join '') -match '^[-\s]+$') { continue } # separator row
            if (-not $inTable) { [void]$html.Append("<table>"); $inTable = $true; $isHeader = $true }
            $tag = if ($isHeader) { 'th' } else { 'td' }
            [void]$html.Append("<tr>")
            foreach ($c in $cells) { [void]$html.Append("<$tag>$(Inline $c)</$tag>") }
            [void]$html.Append("</tr>")
            $isHeader = $false
            continue
        } elseif ($inTable) { [void]$html.Append("</table>"); $inTable = $false }

        if ($ln -match '^### (.+)') { [void]$html.Append("<h3>$(Inline $Matches[1])</h3>") }
        elseif ($ln -match '^## (.+)') { [void]$html.Append("<h2>$(Inline $Matches[1])</h2>") }
        elseif ($ln -match '^# (.+)')  { [void]$html.Append("<h1>$(Inline $Matches[1])</h1>") }
        elseif ($ln -match '^> (.+)')  { [void]$html.Append("<blockquote>$(Inline $Matches[1])</blockquote>") }
        elseif ($ln -match '^- (.+)')  {
            if (-not $inList) { [void]$html.Append("<ul>"); $inList = $true }
            [void]$html.Append("<li>$(Inline $Matches[1])</li>")
        }
        elseif ($ln -match '^---\s*$') { [void]$html.Append("<hr>") }
        elseif ($ln.Trim() -eq '') { if ($inList) { [void]$html.Append("</ul>"); $inList=$false } }
        else {
            if ($inList) { [void]$html.Append("</ul>"); $inList=$false }
            [void]$html.Append("<p>$(Inline $ln)</p>")
        }
    }
    if ($inTable) { [void]$html.Append("</table>") }
    if ($inList)  { [void]$html.Append("</ul>") }
    return $html.ToString()
}

function Build-Html {
    param([string]$Md, [string]$Title, [string]$TemplateDir)
    $tpl = Read-Utf8 (Join-Path $TemplateDir "report.html.tmpl")
    $content = Convert-MarkdownToHtml $Md
    # literal replacement (content/title may contain regex metacharacters)
    return $tpl.Replace('{{TITLE}}', $Title).Replace('{{CONTENT}}', $content)
}

function Build-Annex {
    param([pscustomobject]$Facts)
    $parts = @("## Anexo - detalhamento")
    if ($Facts.serviceMap.backup) {
        $failRows = @($Facts.backup.perMachine | Where-Object { $_.err -gt 0 } | ForEach-Object { ,@($_.machine, $_.runs, $_.err) })
        if ($failRows.Count) { $parts += "`n### Execuções com falha por máquina`n" + (New-MdTable @("Máquina","Execuções","Falhas") $failRows) }
        $agentRows = @($Facts.lifecycle.agents | ForEach-Object { ,@($_.name, $_.os, $_.coreVersion, $(if($_.online){"online"}else{"offline"})) })
        if ($agentRows.Count) { $parts += "`n### Inventário de agentes`n" + (New-MdTable @("Nome","SO","Versão","Estado") $agentRows) }
    }
    if ($Facts.serviceMap.edr -and $Facts.security) {
        $typeRows = @($Facts.security.byType | ForEach-Object { ,@($_.type, $_.count) })
        if ($typeRows.Count) { $parts += "`n### Alertas por tipo`n" + (New-MdTable @("Tipo","Qtd") $typeRows) }
    }
    if ($Facts.audit -and -not $Facts.audit.accessPending -and @($Facts.audit.events).Count) {
        $aRows = @($Facts.audit.events | ForEach-Object { ,@($_.date, $_.initiator, $_.action, $_.object) })
        $parts += "`n### Eventos de auditoria`n" + (New-MdTable @("Data","Autor","Ação","Objeto") $aRows)
    }
    ($parts -join "`n")
}

function Build-Report {
    param([pscustomobject]$Facts)
    $cover = @"
# Relatório Mensal MSS - Acronis Cyber Protect Cloud
## Cliente: $($Facts.meta.client) - $($Facts.meta.month)
**Provedor:** Contego Security · **Data center:** $($Facts.meta.dataCenter) · **Gerado:** $($Facts.meta.generatedAt) (UTC)

---

"@
    $sections = @(
        Build-ExecSummary $Facts
        Build-Environment $Facts
        Build-Backup $Facts
        Build-Incidents $Facts
        Build-Security $Facts
        Build-Va $Facts
        Build-M365 $Facts
        Build-Audit $Facts
        Build-Contract $Facts
        Build-ActionPlan $Facts
        Build-Methodology $Facts
        Build-Annex $Facts
    ) | Where-Object { $_ -and $_.Trim() }
    $cover + ($sections -join "`n`n")
}

function Convert-HtmlToPdf {
    param([string]$HtmlPath, [string]$PdfPath)
    $have = { param($n) [bool](Get-Command $n -ErrorAction SilentlyContinue) }
    if (& $have 'wkhtmltopdf') {
        & wkhtmltopdf --quiet $HtmlPath $PdfPath 2>$null
        if (Test-Path $PdfPath) { return [pscustomobject]@{ pdf=$true; tool='wkhtmltopdf' } }
    }
    if (& $have 'pandoc') {
        & pandoc $HtmlPath -o $PdfPath 2>$null
        if (Test-Path $PdfPath) { return [pscustomobject]@{ pdf=$true; tool='pandoc' } }
    }
    $edge = @("$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
              "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") |
              Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($edge) {
        $uri = ([Uri]("file://" + ($HtmlPath -replace '\\','/'))).AbsoluteUri
        # redirect Edge/Chromium stdout+stderr to temp files so its internal logging doesn't clutter the console
        $errLog = [System.IO.Path]::GetTempFileName()
        $outLog = [System.IO.Path]::GetTempFileName()
        $proc = Start-Process -FilePath $edge `
            -ArgumentList "--headless","--disable-gpu","--disable-logging","--log-level=3","--no-pdf-header-footer","--print-to-pdf=$PdfPath",$uri `
            -PassThru -NoNewWindow -RedirectStandardError $errLog -RedirectStandardOutput $outLog
        $proc.WaitForExit(30000) | Out-Null
        if (-not $proc.HasExited) { try { $proc.Kill() } catch {} }
        Remove-Item $errLog, $outLog -ErrorAction SilentlyContinue
        if (Test-Path $PdfPath) { return [pscustomobject]@{ pdf=$true; tool='edge' } }
    }
    return [pscustomobject]@{ pdf=$false; tool='none' }
}

function Build-GammaPrompt {
    param([pscustomobject]$Facts)
    $sb=[System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("Crie uma apresentação cliente-facing (PT-BR), padrão MSSP, da Contego Security (cor da marca #602E91) sobre Acronis Cyber Protect Cloud.")
    [void]$sb.AppendLine("CLIENTE: $($Facts.meta.client)  PERÍODO: $($Facts.meta.month)")
    [void]$sb.AppendLine("Slide 1 (capa): marca Contego, cliente, período, postura $($Facts.analysis.healthScore)/100.")
    [void]$sb.AppendLine("Slide 2 (dashboard executivo): KPIs e gráficos:")
    if ($Facts.serviceMap.backup -and $Facts.backup) { [void]$sb.AppendLine("- Backup: $($Facts.backup.successPct)% sucesso em $($Facts.backup.runs) execuções (donut ok/erro).") }
    if ($Facts.storage -and $Facts.storage.occupancyPct) { [void]$sb.AppendLine("- Armazenamento: $($Facts.storage.occupancyPct)% do contratado (medidor).") }
    if ($Facts.serviceMap.edr -and $Facts.security) { [void]$sb.AppendLine("- Segurança/EDR: $($Facts.security.total) alertas, $($Facts.security.bySeverity.critical) críticos (barras por severidade).") }
    if ($Facts.lifecycle) { [void]$sb.AppendLine("- Frota: $($Facts.lifecycle.total) agentes, $($Facts.lifecycle.online) online.") }
    if ($Facts.serviceMap.backup) { [void]$sb.AppendLine("Slide: Atividade de backup e pontos críticos (agrupados por causa).") }
    if ($Facts.serviceMap.edr)    { [void]$sb.AppendLine("Slide: Postura de segurança e higiene operacional (alertas por tema).") }
    if ($Facts.serviceMap.va)     { [void]$sb.AppendLine("Slide: Vulnerabilidades ($($Facts.vulnerability.scansRun) varreduras).") }
    if ($Facts.serviceMap.m365)   { [void]$sb.AppendLine("Slide: Proteção Microsoft 365.") }
    if ($Facts.serviceUsage -and @($Facts.serviceUsage).Count -gt 0) { [void]$sb.AppendLine("Slide: Serviços habilitados (cota vs uso).") }
    [void]$sb.AppendLine("Slide: Plano de ação (priorizado, agrupado por causa).")
    [void]$sb.AppendLine("Não invente números além dos fornecidos.")
    $sb.ToString()
}

Export-ModuleMember -Function Read-Utf8, Write-Utf8, New-MdTable, Get-StatusWord, Build-ExecSummary, Build-Environment, Build-Backup, Build-Incidents, Build-Security, Build-Va, Build-M365, Build-Audit, Build-Contract, Build-ActionPlan, Build-Methodology, Build-Annex, Build-Report, Convert-MarkdownToHtml, Build-Html, Convert-HtmlToPdf, Build-GammaPrompt