$ErrorActionPreference = "Stop"
function Read-Utf8([string]$p) { [System.IO.File]::ReadAllText($p, [System.Text.UTF8Encoding]::new($false)) }
function Write-Utf8([string]$p,[string]$t) { [System.IO.File]::WriteAllText($p, $t, (New-Object System.Text.UTF8Encoding($false))) }

function New-SvgDonut {
    param([object[]]$Segments,[int]$Size=160)
    $r=60; $c=[math]::PI*2*$r; $cx=$Size/2; $cy=$Size/2
    $total=(@($Segments) | ForEach-Object { [double]$_.value } | Measure-Object -Sum).Sum
    $sb=[System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $Size $Size' width='$Size' height='$Size'>")
    if (-not $total) {
        [void]$sb.Append("<circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='#e5e7eb' stroke-width='20'/>")
        [void]$sb.Append("<text x='$cx' y='$cy' text-anchor='middle' dy='.3em' font-size='12' fill='#6b7280'>sem dados</text></svg>")
        return $sb.ToString()
    }
    $offset=0
    foreach ($s in $Segments) {
        $frac=[double]$s.value/$total; $len=$frac*$c; $gap=$c-$len
        [void]$sb.Append("<circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$($s.color)' stroke-width='20' stroke-dasharray='$([math]::Round($len,2)) $([math]::Round($gap,2))' stroke-dashoffset='$([math]::Round(-$offset,2))' transform='rotate(-90 $cx $cy)'/>")
        $offset+=$len
    }
    [void]$sb.Append("</svg>")
    $sb.ToString()
}

function New-SvgBars {
    param([object[]]$Bars,[int]$Width=320,[int]$Height=160)
    $bars=@($Bars); $sb=[System.Text.StringBuilder]::new()
    [void]$sb.Append("<svg viewBox='0 0 $Width $Height' width='$Width' height='$Height'>")
    $max=(@($bars) | ForEach-Object { [double]$_.value } | Measure-Object -Maximum).Maximum
    if (-not $max) { [void]$sb.Append("<text x='8' y='20' font-size='12' fill='#6b7280'>sem dados</text></svg>"); return $sb.ToString() }
    $n=$bars.Count; $bw=[math]::Floor(($Width-20)/[math]::Max($n,1))-8
    for ($i=0;$i -lt $n;$i++) {
        $h=[int](($Height-30)*([double]$bars[$i].value/$max)); $x=10+$i*($bw+8); $y=$Height-20-$h
        [void]$sb.Append("<rect x='$x' y='$y' width='$bw' height='$h' fill='$($bars[$i].color)' rx='3'/>")
        [void]$sb.Append("<text x='$($x+$bw/2)' y='$($Height-6)' text-anchor='middle' font-size='9' fill='#353535'>$($bars[$i].label)</text>")
        [void]$sb.Append("<text x='$($x+$bw/2)' y='$($y-3)' text-anchor='middle' font-size='9' fill='#353535'>$($bars[$i].value)</text>")
    }
    [void]$sb.Append("</svg>"); $sb.ToString()
}

function New-SvgGauge {
    param([double]$Pct,[string]$Color='#602E91',[int]$Size=160)
    $p=[math]::Max(0,[math]::Min(100,$Pct)); $r=60; $c=[math]::PI*2*$r; $cx=$Size/2; $cy=$Size/2
    $len=($p/100)*$c; $gap=$c-$len
    "<svg viewBox='0 0 $Size $Size' width='$Size' height='$Size'>" +
    "<circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='#e5e7eb' stroke-width='16'/>" +
    "<circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$Color' stroke-width='16' stroke-dasharray='$([math]::Round($len,2)) $([math]::Round($gap,2))' transform='rotate(-90 $cx $cy)'/>" +
    "<text x='$cx' y='$cy' text-anchor='middle' dy='.3em' font-size='22' font-weight='bold' fill='#353535'>$([int]$Pct)%</text></svg>"
}

function New-Legend {
    param([object[]]$Segments)
    $items = @($Segments | Where-Object { $_ } | ForEach-Object { "<span><i style='background:$($_.color)'></i>$($_.label)</span>" })
    if (-not $items.Count) { return "" }
    "<div class='legend'>$($items -join '')</div>"
}

function New-SvgSparkline {
    param([double[]]$Values,[int]$Width=160,[int]$Height=40)
    $v=@($Values); if ($v.Count -lt 2) { return "<svg width='$Width' height='$Height'><text x='4' y='24' font-size='10' fill='#6b7280'>—</text></svg>" }
    $min=($v|Measure-Object -Minimum).Minimum; $max=($v|Measure-Object -Maximum).Maximum; $rng=[math]::Max($max-$min,1)
    $step=$Width/([math]::Max($v.Count-1,1)); $pts=@()
    for ($i=0;$i -lt $v.Count;$i++){ $x=[int]($i*$step); $y=[int]($Height-2-(($v[$i]-$min)/$rng)*($Height-4)); $pts+="$x,$y" }
    "<svg viewBox='0 0 $Width $Height' width='$Width' height='$Height'><polyline fill='none' stroke='#602E91' stroke-width='2' points='$($pts -join " ")'/></svg>"
}

function Get-EmbeddedAsset {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return '' }
    $ext = [System.IO.Path]::GetExtension($Path).TrimStart('.').ToLower()
    $mime = switch ($ext) { 'png'{'image/png'} 'jpg'{'image/jpeg'} 'jpeg'{'image/jpeg'} 'svg'{'image/svg+xml'} default{'application/octet-stream'} }
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($Path))
    "data:$mime;base64,$b64"
}
function New-KpiCard {
    param([string]$Label,[string]$Value,[string]$Status='ok')
    $color = switch ($Status) { 'warn'{'#e8a100'} 'crit'{'#c5221f'} default{'#1e8e3e'} }
    "<div class='kpi'><div class='kpi-v' style='color:$color'>$Value</div><div class='kpi-l'>$Label</div></div>"
}
function Get-Posture {
    param([int]$Score)
    if ($Score -ge 85) { [pscustomobject]@{ band='Saudável'; color='#1e8e3e' } }
    elseif ($Score -ge 60) { [pscustomobject]@{ band='Atenção'; color='#e8a100' } }
    else { [pscustomobject]@{ band='Crítico'; color='#c5221f' } }
}

$script:AlertCatalog = $null
function Get-AlertCatalog {
    <# Loads the enriched alert catalog (bucket/severity/labelPt/descricao/causa/acao/doc), cached. #>
    if ($null -eq $script:AlertCatalog) {
        $p = Join-Path $PSScriptRoot "..\config\alert_types_pt.json"
        $script:AlertCatalog = if (Test-Path $p) { [System.IO.File]::ReadAllText($p,[System.Text.UTF8Encoding]::new($false)) | ConvertFrom-Json } else { [pscustomobject]@{} }
    }
    $script:AlertCatalog
}

function Build-AlertGlossary {
    <# Analyst glossary for the alert types present in this client's security.byType. Inner HTML only
       (caller wraps with the numbered section heading). Returns "" when there are no alerts. #>
    param([pscustomobject]$Facts)
    if (-not ($Facts.serviceMap.edr -and $Facts.security)) { return "" }
    $types = @($Facts.security.byType | ForEach-Object { "$($_.type)" } | Where-Object { $_ } | Select-Object -Unique)
    if (-not $types.Count) { return "" }
    $cat = Get-AlertCatalog
    $esc = { param($s) "$s" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }
    $order = @{ 'critica'=0; 'alta'=1; 'media'=2; 'baixa'=3; 'informativa'=4 }
    $rows = foreach ($t in $types) {
        $e = $cat.PSObject.Properties[$t]
        if ($e) { $v=$e.Value; [pscustomobject]@{ label=$v.labelPt; sev=$v.severity; desc=$v.descricao; causa=$v.causa; acao=$v.acao; ord=$(if($order.ContainsKey("$($v.severity)")){$order["$($v.severity)"]}else{5}) } }
        else    { [pscustomobject]@{ label=$t; sev='-'; desc='(sem descrição catalogada)'; causa='-'; acao='-'; ord=5 } }
    }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<p><small>Refer&ecirc;ncia para analistas: significado, causa prov&aacute;vel e a&ccedil;&atilde;o de cada alerta presente neste relat&oacute;rio.</small></p>")
    [void]$sb.Append("<table><tr><th>Alerta</th><th>Sev.</th><th>O que significa</th><th>Causa prov&aacute;vel</th><th>A&ccedil;&atilde;o recomendada</th></tr>")
    foreach ($r in ($rows | Sort-Object ord, label)) {
        [void]$sb.Append("<tr><td>$(& $esc $r.label)</td><td>$(& $esc $r.sev)</td><td>$(& $esc $r.desc)</td><td>$(& $esc $r.causa)</td><td>$(& $esc $r.acao)</td></tr>")
    }
    [void]$sb.Append("</table>")
    $sb.ToString()
}

function Get-Anomalies {
    <# Flag what is NOT normal for a client: total/high backup failure, critical alerts,
       over-quota services, heavy agent offline, near-full storage. Returns PT-BR strings. #>
    param([pscustomobject]$Facts)
    # Plain accented PT-BR (module is UTF-8 BOM); the renderer escapes via & $esc, so do NOT use HTML entities here.
    $a = @()
    if ($Facts.serviceMap.backup -and $Facts.backup -and [int]$Facts.backup.runs -gt 0) {
        $sp = [double]$Facts.backup.successPct
        if ($sp -eq 0)      { $a += "Falha total de backup: 0% de sucesso em $($Facts.backup.runs) execuções no período." }
        elseif ($sp -lt 50) { $a += "Backup com alta taxa de falha: apenas $sp% de sucesso em $($Facts.backup.runs) execuções." }
    }
    if ($Facts.serviceMap.edr -and $Facts.security -and [int]$Facts.security.bySeverity.critical -gt 0) {
        $a += "$($Facts.security.bySeverity.critical) alerta(s) crítico(s) de segurança em aberto."
    }
    foreach ($s in @($Facts.serviceUsage)) {
        if ($null -ne $s.pct -and [int]$s.pct -gt 100) { $a += "$($s.labelPt): consumo acima da cota contratada ($($s.pct)%)." }
    }
    if ($Facts.lifecycle -and [int]$Facts.lifecycle.total -gt 0) {
        $offl = @($Facts.lifecycle.offline).Count
        if ($offl -gt 0 -and ($offl / [double]$Facts.lifecycle.total) -ge 0.2) { $a += "$offl de $($Facts.lifecycle.total) agentes offline (>20% da frota)." }
    }
    if ($Facts.storage -and [double]$Facts.storage.occupancyPct -ge 90) {
        $a += "Armazenamento em $($Facts.storage.occupancyPct)% da cota contratada."
    }
    @($a)
}

function Get-EffectiveHealth {
    param([pscustomobject]$Facts)
    $eff = [int]$Facts.analysis.healthScore
    $hasAct = [bool]($Facts.serviceMap.backup -and $Facts.backup -and [int]$Facts.backup.runs -gt 0)
    if ($Facts.serviceMap.backup -and -not $hasAct) {
        $sp = if ($Facts.backup) { [double]$Facts.backup.successPct } else { 0 }
        # mirrors Get-Analysis backup-success weight (0.5 * (100 - successPct))
        $eff += [int]([math]::Round(0.5 * (100 - $sp)))
    }
    if ($eff -gt 100) { $eff = 100 }
    if ($eff -lt 0)  { $eff = 0 }
    return $eff
}


function ConvertTo-PtPattern {
    param([string]$Pattern)
    switch ($Pattern) { '100% failure' {'100% de falha'} 'high failure count' {'alto número de falhas'} default {$Pattern} }
}
function Get-GroupedActions {
    param([pscustomobject]$Facts)
    $acts=@()
    # backup incidents grouped by cause
    if ($Facts.serviceMap.backup -and @($Facts.incidents).Count) {
        $Facts.incidents | Group-Object cause | Sort-Object Count -Descending | ForEach-Object {
            $acts += "Tratar causa raiz: $($_.Name) — $($_.Count) máquina(s) afetada(s)."
        }
    }
    # capacity
    if ($Facts.storage -and $Facts.storage.occupancyPct -ge 80) { $acts += "Armazenamento em $($Facts.storage.occupancyPct)% do contratado — avaliar expansão." }
    # outdated / offline
    if ($Facts.lifecycle) {
        if (@($Facts.lifecycle.outdated).Count) { $acts += "Atualizar $(@($Facts.lifecycle.outdated).Count) agente(s) desatualizado(s)." }
        if (@($Facts.lifecycle.offline).Count)  { $acts += "Investigar $(@($Facts.lifecycle.offline).Count) agente(s) offline." }
    }
    # security buckets (top 3 hygiene + posture)
    if ($Facts.serviceMap.edr -and $Facts.security) {
        $b = Get-AlertBuckets $Facts.security
        @($b.posture | Sort-Object count -Descending | Select-Object -First 2) | ForEach-Object { $acts += "Segurança: $($_.labelPt) ($($_.count))." }
        @($b.hygiene | Sort-Object count -Descending | Select-Object -First 3) | ForEach-Object { $acts += "Higiene: $($_.labelPt) ($($_.count))." }
    }
    if (-not $acts.Count) { $acts += "Nenhuma ação crítica no período." }
    @($acts)
}
function Get-Narrative {
    param([pscustomobject]$Facts)
    $eff = Get-EffectiveHealth $Facts
    $p = Get-Posture $eff
    $parts = @("No período, a postura geral do ambiente foi avaliada como **$($p.band.ToLower())** ($eff/100).")
    $hasAct = [bool]($Facts.serviceMap.backup -and $Facts.backup -and [int]$Facts.backup.runs -gt 0)
    if ($hasAct) {
        $parts += "Backup: $($Facts.backup.successPct)% de sucesso em $($Facts.backup.runs) execuções."
        if (@($Facts.incidents).Count) { $parts += "Destaque crítico: $(@($Facts.incidents).Count) máquina(s) com falhas recorrentes." }
    } elseif ($Facts.serviceMap.backup) {
        $parts += "Backup habilitado, sem atividade no período neste tenant."
    }
    if ($Facts.serviceMap.edr -and $Facts.security) { $parts += "Segurança: $($Facts.security.total) alertas ($($Facts.security.bySeverity.critical) críticos)." }
    if ($Facts.lifecycle) { $parts += "Frota: $($Facts.lifecycle.total) agentes, $($Facts.lifecycle.online) online." }
    ($parts -join ' ')
}
function Get-SlaCompliance {
    param([pscustomobject]$Facts,[pscustomobject]$Sla)
    $rows=@()
    if ($Facts.serviceMap.backup -and $Facts.backup -and [int]$Facts.backup.runs -gt 0) {
        $rows += [pscustomobject]@{ metric='Sucesso de backup'; target="$($Sla.backupSuccessPct)%"; actual="$($Facts.backup.successPct)%"; met=([double]$Facts.backup.successPct -ge [double]$Sla.backupSuccessPct) }
    }
    if ($Facts.lifecycle -and $Facts.lifecycle.total) {
        $onlinePct=[int](100.0*$Facts.lifecycle.online/$Facts.lifecycle.total)
        $rows += [pscustomobject]@{ metric='Agentes online'; target="$($Sla.agentsOnlinePct)%"; actual="$onlinePct%"; met=($onlinePct -ge [double]$Sla.agentsOnlinePct) }
    }
    if ($Facts.storage -and $Facts.storage.occupancyPct) {
        $rows += [pscustomobject]@{ metric='Ocupação de armazenamento'; target="<= $($Sla.storageOccupancyMaxPct)%"; actual="$($Facts.storage.occupancyPct)%"; met=([double]$Facts.storage.occupancyPct -le [double]$Sla.storageOccupancyMaxPct) }
    }
    @($rows)
}
function Build-OutOfContract {
    <# Inner HTML for the "Pontos de atenção — fora do contrato" section. "" when empty.
       The numbered heading is added by the caller via Sec(). #>
    param([pscustomobject]$Facts)
    if (-not $Facts.outOfContract -or @($Facts.outOfContract).Count -eq 0) { return "" }
    $items = @($Facts.outOfContract)
    $esc = { param($s) "$s" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($o in $items) {
        [void]$sb.Append("<p><b>$(& $esc $o.labelPt)</b> &mdash; $(& $esc $o.message)</p>")
        $m = $o.metrics
        switch ($o.service) {
            'backup' {
                [void]$sb.Append("<p>$($m.runs) execu&ccedil;&otilde;es &middot; $($m.error) falhas &middot; $($m.successPct)% sucesso.")
                if (@($m.topErrors).Count) {
                    $errs = (@($m.topErrors) | ForEach-Object { "$(& $esc $_.labelPt) ($($_.count))" }) -join '; '
                    [void]$sb.Append(" Principais erros: $errs.")
                }
                [void]$sb.Append("</p>")
            }
            'edr' {
                [void]$sb.Append("<p>$($m.total) alertas.")
                if (@($m.topTypes).Count) { $t=(@($m.topTypes)|ForEach-Object{"$(& $esc $_.labelPt) ($($_.count))"}) -join '; '; [void]$sb.Append(" Tipos principais: $t.") }
                [void]$sb.Append("</p>")
            }
            'va'   { [void]$sb.Append("<p>$($m.scansRun) varreduras de vulnerabilidade.</p>") }
            'm365' { [void]$sb.Append("<p>$($m.seatsUsed) assentos &middot; $($m.storageTB) TB de backup M365.</p>") }
        }
    }
    return $sb.ToString()
}
function Build-HtmlReport {
    param([pscustomobject]$Facts,[object[]]$History,[pscustomobject]$Sla,[string]$ConfigDir)
    $esc = { param($s) "$s" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }
    $logo = Get-EmbeddedAsset (Join-Path $ConfigDir 'contego_logo.jpg')
    $effHealth = Get-EffectiveHealth $Facts
    $pst  = Get-Posture $effHealth
    $hasBackupActivity = [bool]($Facts.serviceMap.backup -and $Facts.backup -and [int]$Facts.backup.runs -gt 0)
    $css = @"
<style>
body{font-family:Segoe UI,Arial,sans-serif;color:#353535;margin:0}
.wrap{max-width:900px;margin:0 auto;padding:32px}
.cover{background:#602E91;color:#fff;padding:36px;border-radius:10px;display:flex;justify-content:space-between;align-items:center}
.cover img{height:54px;background:#fff;padding:6px 10px;border-radius:6px}
.badge{font-size:34px;font-weight:bold;background:#fff;color:$($pst.color);border-radius:10px;padding:10px 18px}
h2{color:#602E91;border-bottom:2px solid #ece7f3;padding-bottom:6px;margin-top:30px}
table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #d8d2e6;padding:7px 10px;text-align:left;font-size:13px}th{background:#602E91;color:#fff}
.kpis{display:flex;gap:14px;flex-wrap:wrap;margin:14px 0}.kpi{flex:1;min-width:130px;border:1px solid #e5e1ef;border-radius:8px;padding:14px;text-align:center}.kpi-v{font-size:26px;font-weight:bold}.kpi-l{font-size:12px;color:#6b7280}
.charts{display:flex;gap:18px;flex-wrap:wrap;align-items:center;margin:14px 0}
blockquote{background:#f6f3fb;border-left:4px solid #602E91;padding:8px 14px;margin:12px 0}
.met{color:#1e8e3e;font-weight:bold}.miss{color:#c5221f;font-weight:bold}
small{color:#6b7280}
.legend{display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:#6b7280;margin-top:4px}.legend span{display:inline-flex;align-items:center;gap:4px}.legend i{width:10px;height:10px;border-radius:2px;display:inline-block}
.anomaly{background:#fdecea;border-left:5px solid #c5221f;padding:10px 16px;margin:16px 0;border-radius:6px}.anomaly strong{color:#c5221f}.anomaly ul{margin:6px 0 0 0;padding-left:18px}
</style>
"@
    $sb=[System.Text.StringBuilder]::new()
    [void]$sb.Append("<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'><title>$(& $esc $Facts.meta.client) - $($Facts.meta.month)</title>$css</head><body><div class='wrap'>")
    # cover
    $logoTag = if ($logo) { "<img src='$logo' alt='Contego Security'/>" } else { "<div style='font-weight:bold;font-size:22px'>CONTEGO SECURITY</div>" }
    [void]$sb.Append("<div class='cover'><div>$logoTag<h1 style='margin:14px 0 4px'>Relat&oacute;rio Mensal de Prote&ccedil;&atilde;o Gerenciada</h1><div>$(& $esc $Facts.meta.client) &middot; $($Facts.meta.month) &middot; $($Facts.meta.dataCenter)</div></div><div style='text-align:center'><div class='badge'>$effHealth/100</div><div style='margin-top:6px'>$($pst.band)</div></div></div>")
    # anomalies banner — flag what's not normal, up top
    $anoms = Get-Anomalies $Facts
    if (@($anoms).Count) {
        [void]$sb.Append("<div class='anomaly'><strong>&#9888; Pontos fora do normal</strong><ul>")
        foreach ($x in $anoms) { [void]$sb.Append("<li>$(& $esc $x)</li>") }
        [void]$sb.Append("</ul></div>")
    }
    $script:n=0; function Sec($t){ $script:n++; "<h2>$script:n. $t</h2>" }
    # dashboard
    [void]$sb.Append((Sec 'Resumo executivo'))
    $kpis=@()
    if ($hasBackupActivity) { $kpis += (New-KpiCard 'Sucesso de backup' "$($Facts.backup.successPct)%" $(if($Facts.backup.successPct -ge 90){'ok'}elseif($Facts.backup.successPct -ge 70){'warn'}else{'crit'})) }
    if ($Facts.storage -and $Facts.storage.occupancyPct) { $kpis += (New-KpiCard 'Armazenamento' "$($Facts.storage.occupancyPct)%" $(if($Facts.storage.occupancyPct -lt 85){'ok'}else{'warn'})) }
    if ($Facts.serviceMap.edr -and $Facts.security) { $kpis += (New-KpiCard 'Alertas cr&iacute;ticos' "$($Facts.security.bySeverity.critical)" $(if($Facts.security.bySeverity.critical -eq 0){'ok'}else{'crit'})) }
    if ($Facts.lifecycle) { $kpis += (New-KpiCard 'Agentes online' "$($Facts.lifecycle.online)/$($Facts.lifecycle.total)" 'ok') }
    [void]$sb.Append("<div class='kpis'>$($kpis -join '')</div>")
    # charts
    $charts=@()
    if ($hasBackupActivity) {
        $seg = @(@{value=$Facts.backup.ok;color='#1e8e3e';label='ok'},@{value=$Facts.backup.warning;color='#e8a100';label='aviso'},@{value=$Facts.backup.error;color='#c5221f';label='erro'})
        $charts += "<div>$(New-SvgDonut $seg)$(New-Legend $seg)</div>"
    }
    if ($Facts.storage -and $Facts.storage.occupancyPct) { $charts += (New-SvgGauge ([double]$Facts.storage.occupancyPct)) }
    if ($Facts.lifecycle -and $Facts.lifecycle.total) {
        $seg = @(@{value=$Facts.lifecycle.online;color='#1e8e3e';label='online'},@{value=($Facts.lifecycle.total-$Facts.lifecycle.online);color='#9ca3af';label='offline'})
        $charts += "<div>$(New-SvgDonut $seg)$(New-Legend $seg)</div>"
    }
    if ($Facts.serviceMap.edr -and $Facts.security) {
        $seg = @(@{value=$Facts.security.bySeverity.critical;color='#c5221f';label='crítico'},@{value=$Facts.security.bySeverity.error;color='#e8a100';label='erro'},@{value=$Facts.security.bySeverity.warning;color='#602E91';label='aviso'})
        $charts += "<div>$(New-SvgBars $seg)$(New-Legend $seg)</div>"
    }
    [void]$sb.Append("<div class='charts'>$($charts -join '')</div>")
    $narr = (Get-Narrative $Facts) -replace '\*\*(.+?)\*\*','<strong>$1</strong>'
    [void]$sb.Append("<blockquote>$narr</blockquote>")
    # quota table (filtered to contracted categories; device/other always shown)
    $allowedCats = @('device','other')
    if ($Facts.serviceMap.backup) { $allowedCats += 'backup' }
    if ($Facts.serviceMap.edr)    { $allowedCats += 'edr' }
    if ($Facts.serviceMap.va)     { $allowedCats += 'va' }
    if ($Facts.serviceMap.m365)   { $allowedCats += 'm365' }
    $visibleUsage = @($Facts.serviceUsage | Where-Object { -not $_.category -or ($allowedCats -contains $_.category) })
    if (@($visibleUsage).Count -gt 0) {
        [void]$sb.Append((Sec 'Servi&ccedil;os habilitados'))
        [void]$sb.Append("<table><tr><th>Servi&ccedil;o</th><th>Cota</th><th>Em uso</th><th>%</th><th>Situa&ccedil;&atilde;o</th></tr>")
        $grouped = @($visibleUsage) | Group-Object labelPt | ForEach-Object {
            $g = $_.Group
            $used = ($g | ForEach-Object { [double]$_.used } | Measure-Object -Sum).Sum
            $quotas = @($g | ForEach-Object { $_.quota } | Where-Object { $_ })
            $quota = if ($quotas.Count) { ($quotas | ForEach-Object { [double]$_ } | Measure-Object -Sum).Sum } else { $null }
            $pct = if ($quota) { [int]([math]::Round(100.0*$used/$quota)) } else { $null }
            [pscustomobject]@{ labelPt=$_.Name; used=$used; quota=$quota; pct=$pct; status=$(if($pct -and $pct -ge 90){'warn'}else{'ok'}) }
        }
        foreach ($s in $grouped) {
            $st = if ($s.status -eq 'warn') { '&#9888;' } else { '&#10003;' }
            [void]$sb.Append("<tr><td>$(& $esc $s.labelPt)</td><td>$(if($s.quota){$s.quota}else{'&mdash;'})</td><td>$($s.used)</td><td>$(if($null -ne $s.pct){"$($s.pct)%"}else{'&mdash;'})</td><td>$st</td></tr>")
        }
        [void]$sb.Append("</table>")
    }
    # SLA
    $sla = Get-SlaCompliance $Facts $Sla
    if (@($sla).Count) {
        [void]$sb.Append((Sec 'SLA e conformidade'))
        [void]$sb.Append("<table><tr><th>M&eacute;trica</th><th>Meta</th><th>Atual</th><th>Status</th></tr>")
        foreach ($r in $sla) { $cls=if($r.met){'met'}else{'miss'}; $txt=if($r.met){'&#10004; Cumprido'}else{'&#10006; N&atilde;o cumprido'}; [void]$sb.Append("<tr><td>$($r.metric)</td><td>$($r.target)</td><td>$($r.actual)</td><td class='$cls'>$txt</td></tr>") }
        [void]$sb.Append("</table>")
    }
    # backup sections
    if ($hasBackupActivity) {
        [void]$sb.Append((Sec 'Atividade de backup'))
        $b=$Facts.backup
        [void]$sb.Append("<table><tr><th>Indicador</th><th>Valor</th></tr><tr><td>Execu&ccedil;&otilde;es</td><td>$($b.runs) ($($b.scheduled) agendadas)</td></tr><tr><td>Sucesso</td><td>$($b.ok) ($($b.successPct)%)</td></tr><tr><td>Erros</td><td>$($b.error)</td></tr><tr><td>Volume transferido</td><td>~$($b.transferredGB) GB</td></tr></table>")
        if (@($Facts.incidents).Count) {
            [void]$sb.Append((Sec 'Pontos cr&iacute;ticos'))
            $Facts.incidents | Group-Object cause | Sort-Object Count -Descending | ForEach-Object {
                $sample = (@($_.Group | ForEach-Object { $_.machine } | Select-Object -First 5) -join ', ')
                [void]$sb.Append("<p>&#8226; <b>$($_.Count) m&aacute;quina(s)</b> &mdash; $(ConvertTo-PtPattern $_.Group[0].pattern). Causa: $(& $esc $_.Name). Ex.: $(& $esc $sample)</p>")
            }
        }
    } elseif ($Facts.serviceMap.backup) {
        [void]$sb.Append("<blockquote>Backup habilitado, sem atividade no período neste tenant.</blockquote>")
    }
    # security sections
    if ($Facts.serviceMap.edr -and $Facts.security) {
        $bk = Get-AlertBuckets $Facts.security
        [void]$sb.Append((Sec 'Postura de segurança'))
        if (@($bk.posture).Count) {
            [void]$sb.Append("<table><tr><th>Tipo</th><th>Qtd</th><th>M&aacute;quinas</th></tr>")
            foreach ($e in ($bk.posture | Sort-Object count -Descending)) {
                $m = (@($e.machines)|Select-Object -First 6) -join ', '
                if (-not $m) { $m = '&mdash;' }
                [void]$sb.Append("<tr><td>$(& $esc $e.labelPt)</td><td>$($e.count)</td><td>$(& $esc $m)</td></tr>")
            }
            [void]$sb.Append("</table>")
        } else { [void]$sb.Append("<blockquote>Nenhum alerta de postura de seguran&ccedil;a no per&iacute;odo.</blockquote>") }
        [void]$sb.Append((Sec 'Higiene operacional'))
        if (@($bk.hygiene).Count) {
            [void]$sb.Append("<table><tr><th>Tema</th><th>Alertas</th></tr>")
            foreach ($e in ($bk.hygiene | Sort-Object count -Descending)) { [void]$sb.Append("<tr><td>$(& $esc $e.labelPt)</td><td>$($e.count)</td></tr>") }
            [void]$sb.Append("</table>")
        } else { [void]$sb.Append("<blockquote>Nenhum alerta de higiene operacional no per&iacute;odo.</blockquote>") }
    }
    # VA
    if ($Facts.serviceMap.va -and $Facts.vulnerability) {
        [void]$sb.Append((Sec 'Avalia&ccedil;&atilde;o de vulnerabilidades'))
        [void]$sb.Append("<p>Varreduras no m&ecirc;s: $($Facts.vulnerability.scansRun) ($($Facts.vulnerability.scanFailures) com falha). &Uacute;ltima: $($Facts.vulnerability.lastScanDate).</p>")
    }
    # M365
    if ($Facts.serviceMap.m365) {
        [void]$sb.Append((Sec 'Prote&ccedil;&atilde;o Microsoft 365'))
        $m = $Facts.manual
        if ($m -and $m.m365SeatsUsed) {
            [void]$sb.Append("<table><tr><th>Indicador</th><th>Valor</th></tr>")
            [void]$sb.Append("<tr><td>Seats Microsoft 365 protegidos</td><td>$($m.m365SeatsUsed) de $($m.m365SeatsContracted)</td></tr>")
            if ($m.m365SeatsShared) { [void]$sb.Append("<tr><td>Seats compartilhados</td><td>$($m.m365SeatsShared)</td></tr>") }
            if ($m.m365SharepointSites) { [void]$sb.Append("<tr><td>Sites SharePoint protegidos</td><td>$($m.m365SharepointSites)</td></tr>") }
            if ($m.m365BackupStorageTB) { [void]$sb.Append("<tr><td>Armazenamento de backup M365</td><td>$($m.m365BackupStorageTB) TB</td></tr>") }
            [void]$sb.Append("</table>")
        }
        else { [void]$sb.Append("<blockquote>Dados do console M365 pendentes (tenant dedicado).</blockquote>") }
    }
    # out-of-contract attention section
    $ooc = Build-OutOfContract $Facts
    if ($ooc) {
        [void]$sb.Append((Sec 'Pontos de aten&ccedil;&atilde;o &mdash; fora do contrato'))
        [void]$sb.Append($ooc)
    }
    # audit
    [void]$sb.Append((Sec 'Altera&ccedil;&otilde;es de licen&ccedil;a e configura&ccedil;&atilde;o'))
    if (-not $Facts.audit -or $Facts.audit.accessPending) { [void]$sb.Append("<blockquote>Acesso pendente: conceda o papel de auditoria ao cliente da API para popular esta se&ccedil;&atilde;o.</blockquote>") }
    else { [void]$sb.Append("<table><tr><th>Data</th><th>Autor</th><th>A&ccedil;&atilde;o</th><th>Objeto</th></tr>"); foreach($e in @($Facts.audit.events)){ [void]$sb.Append("<tr><td>$($e.date)</td><td>$(& $esc $e.initiator)</td><td>$(& $esc $e.action)</td><td>$(& $esc $e.object)</td></tr>") }; [void]$sb.Append("</table>") }
    # action plan
    [void]$sb.Append((Sec 'Plano de a&ccedil;&atilde;o'))
    [void]$sb.Append("<ol>"); foreach ($a in Get-GroupedActions $Facts) { [void]$sb.Append("<li>$(& $esc $a)</li>") }; [void]$sb.Append("</ol>")
    # trend
    if ($Facts.serviceMap.backup -and @($History).Count -ge 2) {
        [void]$sb.Append((Sec 'Tend&ecirc;ncia'))
        $succ=@($History | ForEach-Object { [double]$_.successPct })
        [void]$sb.Append("<p>Sucesso de backup (m&ecirc;s a m&ecirc;s): $(New-SvgSparkline $succ)</p>")
    }
    # alert glossary appendix (analyst reference for the alert types present in this report)
    $glossary = Build-AlertGlossary $Facts
    if ($glossary) { [void]$sb.Append((Sec 'Gloss&aacute;rio de alertas')); [void]$sb.Append($glossary) }
    # methodology
    $gateNote = if ($Facts.meta.serviceMapSource -eq 'contract') { 'Se&ccedil;&otilde;es definidas pelo contrato do cliente.' } elseif ($Facts.meta.serviceMapSource -eq 'usages') { 'Se&ccedil;&otilde;es definidas pelas licen&ccedil;as ativas (uso).' } else { '' }
    [void]$sb.Append("<hr><small>M&eacute;todo: coleta somente-leitura da API Acronis; janela $($Facts.meta.windowUtc) (UTC); armazenamento em base bin&aacute;ria; reten&ccedil;&atilde;o do audit log 180 dias. $gateNote Gerado em $($Facts.meta.generatedAt).</small>")
    [void]$sb.Append("</div></body></html>")
    $sb.ToString()
}

Export-ModuleMember -Function Read-Utf8, Write-Utf8, New-SvgDonut, New-SvgBars, New-SvgGauge, New-SvgSparkline, New-Legend, Get-EmbeddedAsset, New-KpiCard, Get-Posture, Get-Anomalies, Get-AlertCatalog, Build-AlertGlossary, Get-EffectiveHealth, ConvertTo-PtPattern, Get-GroupedActions, Get-Narrative, Get-SlaCompliance, Build-OutOfContract, Build-HtmlReport