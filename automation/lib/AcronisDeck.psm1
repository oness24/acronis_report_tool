# automation/lib/AcronisDeck.psm1 — Gamma-style 16:9 client deck (facts -> HTML), gated to contracted services.
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "AcronisReport.psm1") -Force -DisableNameChecking

$script:DeckEsc = { param($s) "$s" -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }
$script:MesesPt = @('','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro')

function _mesPt([string]$ym){
    try { $p=$ym.Split('-'); return "$($script:MesesPt[[int]$p[1]]) $($p[0])" } catch { return $ym }
}

function Get-DeckLogo([string]$ConfigDir){
    if ($ConfigDir) {
        $p = Join-Path $ConfigDir 'contego_logo.jpg'
        if (Test-Path $p) {
            $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($p))
            return "<div class='logo'><img src='data:image/jpeg;base64,$b64' alt='Contego Security'></div>"
        }
    }
    return "<div class='logo' style='color:#622E95;font-family:Poppins,sans-serif;font-weight:700;padding:14px 18px'>Contego Security</div>"
}

function Get-DeckCss {
@"
<style>
@page { size: 1280px 720px; margin: 0; }
*{box-sizing:border-box}
body{margin:0;background:#622E95;font-family:'Inter',system-ui,sans-serif;color:#fff}
.slide{width:1280px;height:720px;position:relative;overflow:hidden;background:#622E95;page-break-after:always}
.slide:last-child{page-break-after:auto}
.pad{position:absolute;inset:0;padding:62px 72px}
h1,h2,h3{font-family:'Poppins',system-ui,sans-serif;margin:0;font-weight:700;letter-spacing:-.01em}
.eyebrow{display:inline-block;font-family:'Poppins',sans-serif;font-size:12.5px;font-weight:600;letter-spacing:.13em;text-transform:uppercase;color:#30D0F0;border:1px solid #30D0F0;border-radius:7px;padding:7px 13px;margin-bottom:22px}
.eyebrow.teal{color:#9ff;background:#0C3B3A;border:none}
.stitle{font-size:46px;font-weight:800;line-height:1.04;margin-bottom:8px}
.sub{color:#D9CBEC;font-size:17px;line-height:1.5;max-width:60ch}
.logo{position:absolute;top:54px;right:72px;background:#fff;border-radius:11px;padding:12px 18px}
.logo img{height:34px;display:block}
.cover h1{font-size:60px;font-weight:800;line-height:1.03}
.cover .right{position:absolute;left:600px;right:72px;top:50%;transform:translateY(-50%)}
.cover .meta{font-family:'Poppins',sans-serif;font-weight:700;font-size:20px;margin:26px 0 30px}
.cover .meta .c{color:#30D0F0}
.cover .desc{color:#D9CBEC;font-size:17px;line-height:1.6;max-width:48ch}
.cover .art{position:absolute;left:-30px;top:120px;width:520px;height:520px}
.cover .art .ln{fill:none;stroke:#fff;stroke-width:7;stroke-linecap:round;stroke-linejoin:round;opacity:.92}
.kgrid{display:grid;grid-template-columns:1fr 1fr;gap:30px 70px;margin-top:30px}
.kpi{text-align:center}
.kpi .n{font-family:'Poppins',sans-serif;font-weight:800;font-size:74px;line-height:1}
.kpi .t{font-family:'Poppins',sans-serif;font-weight:700;font-size:22px;margin-top:8px}
.kpi .c{color:#D9CBEC;font-size:15px;margin-top:7px}
.kpi .c .hl{color:#30D0F0}
.n.gold{color:#F0D030}.n.blue{color:#3445FF}.n.orange{color:#F0A040}.n.cyan{color:#30D0F0}.n.green{color:#3FD07A}
.callout{position:absolute;left:72px;right:72px;bottom:54px;background:#4A3F01;border-radius:12px;padding:20px 26px;color:#F0D030;font-size:16px;line-height:1.5;display:flex;gap:14px;align-items:flex-start}
.callout b{color:#fff}.callout .ic{font-size:20px;line-height:1}
table{width:100%;border-collapse:collapse;margin-top:8px;font-size:17px;border-radius:10px;overflow:hidden}
thead th{background:#3445FF;color:#fff;font-family:'Poppins',sans-serif;font-weight:600;text-align:left;padding:16px 22px;font-size:16px}
tbody td{padding:16px 22px;color:#efe9f7;border-bottom:1px solid rgba(255,255,255,.10)}
tbody tr{background:rgba(255,255,255,.04)}
td .pct{color:#30D0F0;font-family:'Poppins',sans-serif;font-weight:700}
td .pct.warn{color:#F0D030}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:6px}
.bluebox{background:#3445FF;border-radius:14px;padding:26px 28px}
.bluebox .h{font-family:'Poppins',sans-serif;font-weight:700;font-size:22px}
.bluebox p{color:#e4e0ff;font-size:16px;line-height:1.55;margin:10px 0 0}
.ocard{border:1px solid #30D0F0;border-radius:12px;padding:20px 22px;margin-bottom:18px}
.ocard .h{font-family:'Poppins',sans-serif;font-weight:700;font-size:20px}
.ocard p{color:#D9CBEC;font-size:15px;line-height:1.5;margin:8px 0 0}
.ttl-sm{font-family:'Poppins',sans-serif;font-weight:700;font-size:22px;margin-bottom:16px}
.bars{display:flex;align-items:flex-end;gap:46px;height:300px;padding:10px 14px 0}
.bar{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%}
.bar i{width:96px;border-radius:9px 9px 0 0}
.bar .v{font-family:'Poppins',sans-serif;font-weight:800;font-size:30px;margin-bottom:10px}
.bar .l{color:#D9CBEC;font-size:15px;margin-top:12px;font-family:'Poppins',sans-serif;font-weight:600}
.alist{list-style:none;padding:0;margin:6px 0 0}
.alist li{display:flex;gap:16px;align-items:flex-start;padding:15px 0;border-bottom:1px solid rgba(255,255,255,.1)}
.alist .num{font-family:'Poppins',sans-serif;font-weight:800;font-size:26px;color:#30D0F0;width:36px;flex:0 0 auto}
.alist .tx{font-size:17px;line-height:1.4}.alist .tx b{font-family:'Poppins',sans-serif;font-weight:700}
.alist .tag{margin-left:auto;font-family:'Poppins',sans-serif;font-size:12.5px;letter-spacing:.08em;text-transform:uppercase;padding:5px 11px;border-radius:999px;font-weight:600}
.tag.bad{background:rgba(197,34,31,.25);color:#ff9a96}.tag.warn{background:rgba(240,208,48,.18);color:#F0D030}
.foot{position:absolute;left:72px;right:72px;bottom:24px;display:flex;justify-content:space-between;color:#BCA9D6;font-family:'Poppins',sans-serif;font-size:12px;border-top:1px solid rgba(255,255,255,.12);padding-top:12px}
</style>
"@
}

function New-DeckCover($Facts,$ConfigDir){
    $esc=$script:DeckEsc
    $client = & $esc $Facts.meta.client
    $mes = _mesPt $Facts.meta.month
    $sm = $Facts.serviceMap
    $art = @"
<svg class='art' viewBox='0 0 400 400'>
<ellipse class='ln' cx='120' cy='95' rx='58' ry='22'/><ellipse class='ln' cx='70' cy='130' rx='40' ry='16'/>
<rect class='ln' x='150' y='170' width='120' height='95' rx='8'/><path class='ln' d='M150 200 L210 235 L270 200'/>
<rect class='ln' x='190' y='285' width='150' height='100' rx='10'/><circle class='ln' cx='265' cy='335' r='14'/>
<path class='ln' d='M40 250 H120 M60 300 H150 M90 350 H190'/><circle class='ln' cx='40' cy='250' r='6'/><circle class='ln' cx='60' cy='300' r='6'/>
</svg>
"@
    @"
<section class='slide cover'><div class='pad'>
$(Get-DeckLogo $ConfigDir)
$art
<div class='right'>
<h1>Relatório Mensal de<br>Segurança — $client</h1>
<div class='meta'><span class='c'>$mes</span> | Contego Security — Managed Security Services</div>
<p class='desc'>Relatório executivo de monitoramento e proteção gerenciada, elaborado pela equipe de Managed Security Services da Contego Security para os gestores de TI da $client.</p>
</div></div></section>
"@
}

function New-DeckSlide($eyebrow,$eyebrowCls,$title,$inner){
    $ec = if ($eyebrowCls) { " $eyebrowCls" } else { "" }
    @"
<section class='slide'><div class='pad'>
<span class='eyebrow$ec'>$eyebrow</span>
<h2 class='stitle'>$title</h2>
$inner
</div></section>
"@
}

function New-DeckKpiGrid($cards){
    $sb=New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='kgrid'>")
    foreach($c in $cards){
        [void]$sb.Append("<div class='kpi'><div class='n $($c.cls)'>$($c.n)</div><div class='t'>$($c.t)</div><div class='c'>$($c.c)</div></div>")
    }
    [void]$sb.Append("</div>")
    return $sb.ToString()
}

function New-DeckCallout($html){ "<div class='callout'><span class='ic'>&#9888;</span><div>$html</div></div>" }

function Get-DeckVisibleUsage($Facts){
    if (-not $Facts.serviceUsage){ return @() }
    $sm=$Facts.serviceMap; $allowed=@('device','other')
    foreach($k in 'backup','edr','va','m365'){ if ($sm.$k){ $allowed+=$k } }
    $seen=[ordered]@{}
    foreach($r in @($Facts.serviceUsage)){
        $cat=$r.category
        if ($cat -and ($allowed -notcontains $cat)){ continue }
        $lab=$r.labelPt
        if (-not $seen.Contains($lab)){ $seen[$lab]=[pscustomobject]@{labelPt=$lab;used=0.0;quota=$null} }
        $seen[$lab].used += [double]$r.used
        if ($r.quota){ $seen[$lab].quota = [double]([double]$seen[$lab].quota + [double]$r.quota) }
    }
    $out=@()
    foreach($v in $seen.Values){
        $pct = if ($v.quota){ [int][math]::Round(100.0*$v.used/$v.quota) } else { $null }
        $out += [pscustomobject]@{ labelPt=$v.labelPt; used=$v.used; quota=$v.quota; pct=$pct }
    }
    return @($out)
}

function New-DeckServicos($rows){
    $esc=$script:DeckEsc
    $sb=New-Object System.Text.StringBuilder
    [void]$sb.Append("<table><thead><tr><th>Serviço</th><th>Utilizadas</th><th>Cota Total</th><th>Ocupação</th></tr></thead><tbody>")
    $near=$false
    foreach($r in $rows){
        $pcttxt = if ($null -ne $r.pct){ "$($r.pct)%" } else { '&mdash;' }
        $cls = if ($r.pct -and $r.pct -ge 90){ ' warn' } else { '' }
        if ($r.pct -and $r.pct -ge 90){ $near=$true }
        $q = if ($r.quota){ [int]$r.quota } else { '&mdash;' }
        [void]$sb.Append("<tr><td>$(& $esc $r.labelPt)</td><td>$([int]$r.used)</td><td>$q</td><td><span class='pct$cls'>$pcttxt</span></td></tr>")
    }
    [void]$sb.Append("</tbody></table>")
    $inner=$sb.ToString()
    if ($near){ $inner += (New-DeckCallout "<b>Cotas próximas do limite</b> — recomenda-se ajuste contratual caso o parque de máquinas cresça.") }
    return (New-DeckSlide 'Serviços &amp; Cotas' '' 'Utilização de Licenças e Serviços' $inner)
}

function New-DeckResumo($Facts){
    $esc=$script:DeckEsc
    $sec=$Facts.security; $lc=$Facts.lifecycle; $b=$Facts.backup; $st=$Facts.storage; $sm=$Facts.serviceMap
    $accents=@('gold','blue','orange','cyan')
    $tmp=@()
    # (1) Backup %
    if ($b -and [int]$b.runs -gt 0){ $tmp += ,@{n="$($b.successPct)%";t='Sucesso de Backup';c="$($b.runs) execuções"} }
    # (2) Estações Protegidas — device row with largest used, gated on edr or va
    if (($sm.edr -or $sm.va) -and $Facts.serviceUsage){
        $devRows=@($Facts.serviceUsage)|Where-Object{$_.category -eq 'device' -and $_.labelPt -match 'sta'}
        if ($devRows){
            $devBest=$devRows|Sort-Object{[double]$_.used} -Descending|Select-Object -First 1
            if ($devBest -and [double]$devBest.used -gt 0){ $tmp += ,@{n="$([int]$devBest.used)";t='Estações Protegidas';c='Workstations com Gerenciamento Avançado'} }
        }
    }
    # (3) EDR Ativo — edr row used count
    if ($sm.edr -and $Facts.serviceUsage){
        $edrRow=@($Facts.serviceUsage)|Where-Object{$_.category -eq 'edr'}|Select-Object -First 1
        if ($edrRow -and [double]$edrRow.used -gt 0){ $tmp += ,@{n="$([int]$edrRow.used)";t='EDR Ativo';c='Workloads com Segurança Avançada + EDR'} }
    }
    # (4) Agentes Instalados
    if ($lc -and $lc.total){
        $win=(@($lc.byOs)|Where-Object{$_.os -match 'win'}|Select-Object -First 1).count
        $mac=(@($lc.byOs)|Where-Object{$_.os -match 'darwin|mac'}|Select-Object -First 1).count
        $tmp += ,@{n="$($lc.total)";t='Agentes Instalados';c="$([int]$win) Windows + $([int]$mac) macOS"}
    }
    # (5) Alertas
    if ($sec -and [int]$sec.total -gt 0){ $tmp += ,@{n="$($sec.total)";t="Alertas no Período";c="<span class='hl'>$($sec.bySeverity.critical) críticos</span> · $($sec.bySeverity.warning) avisos"} }
    # (6) Armazenamento
    if ($st -and $st.occupancyPct){ $tmp += ,@{n="$($st.occupancyPct)%";t='Armazenamento';c="$($st.cloudGB) de $($st.contractedGB) GB"} }
    $cards = @(); $k=0
    foreach($t in ($tmp | Select-Object -First 4)){ $t.cls=$accents[$k % 4]; $cards += $t; $k++ }
    $grid = New-DeckKpiGrid $cards
    $dest = $null
    $mesNome = (_mesPt $Facts.meta.month).Split(' ')[0].ToLower()
    if ($lc -and @($lc.onboardedThisMonth).Count){ $dest = "<b>Destaque do mês:</b> $(@($lc.onboardedThisMonth).Count) agentes implantados em $mesNome — consolidação da cobertura de proteção na frota." }
    elseif ($Facts.analysis -and @($Facts.analysis.topPriorities).Count){ $dest = "<b>Destaque do mês:</b> $(& $esc $Facts.analysis.topPriorities[0])." }
    $inner = $grid
    if ($dest){ $inner += (New-DeckCallout $dest) }
    return (New-DeckSlide 'Resumo Executivo' 'teal' "Visão Geral — $(_mesPt $Facts.meta.month)" $inner)
}

function New-DeckBackup($Facts){
    $esc=$script:DeckEsc; $b=$Facts.backup
    $cards=@(
        @{n="$($b.successPct)%";cls='green';t='Sucesso de Backup';c="$($b.runs) execuções"},
        @{n="$($b.ok)";cls='cyan';t='Concluídos';c="$($b.error) com erro · $($b.warning) com aviso"},
        @{n="$([int]$b.transferredGB)";cls='blue';t='GB Transferidos';c="no período"}
    )
    $inner = New-DeckKpiGrid $cards
    $top = @($b.errorReasons) | Select-Object -First 1
    if ($top){ $inner += (New-DeckCallout "<b>Principal causa de falha:</b> $(& $esc $top.plainPt) — $($top.count) execução(ões). Recomenda-se correção do destino/credenciais.") }
    return (New-DeckSlide 'Continuidade de Dados' '' 'Atividade de Backup' $inner)
}

function New-DeckSeguranca($Facts){
    $esc=$script:DeckEsc; $sec=$Facts.security; $sev=$sec.bySeverity
    $bk = Get-AlertBuckets $sec
    $maxv = [math]::Max(1, [math]::Max([int]$sev.critical, [math]::Max([int]$sev.warning, [int]$sev.error)))
    $hb={ param($v) [int][math]::Max(4,[math]::Round(100.0*[int]$v/$maxv)) }
    $bars = @"
<div class='bars'>
<div class='bar'><div class='v' style='color:#ff7a74'>$($sev.critical)</div><i style='height:$(& $hb $sev.critical)%;background:#C5221F'></i><div class='l'>Críticos</div></div>
<div class='bar'><div class='v' style='color:#F0D030'>$($sev.warning)</div><i style='height:$(& $hb $sev.warning)%;background:#F0D030'></i><div class='l'>Avisos</div></div>
<div class='bar'><div class='v' style='color:#F0A040'>$($sev.error)</div><i style='height:$(& $hb $sev.error)%;background:#F0A040'></i><div class='l'>Erros</div></div>
</div>
"@
    $items = @(@($bk.posture)+@($bk.hygiene) | Sort-Object count -Descending | Select-Object -First 5)
    $li=New-Object System.Text.StringBuilder
    [void]$li.Append("<div class='ttl-sm'>Principais tipos</div><ul class='alist'>")
    foreach($e in $items){
        $isP = (@($bk.posture).type -contains $e.type)
        $tag = if ($isP){'bad'}else{'warn'}
        [void]$li.Append("<li><span class='tx'><b>$(& $esc $e.labelPt)</b></span><span class='tag $tag'>$($e.count)</span></li>")
    }
    [void]$li.Append("</ul>")
    $inner = "<div class='cols' style='align-items:center'>$bars<div>$($li.ToString())</div></div>"
    return (New-DeckSlide 'Postura de Segurança' '' 'Alertas por Severidade' $inner)
}

function New-DeckVa($Facts){
    $v=$Facts.vulnerability
    $cards=@(@{n="$($v.scansRun)";cls='cyan';t='Varreduras';c='no período'},@{n="$($v.scanFailures)";cls='orange';t='Falhas';c='nas varreduras'})
    return (New-DeckSlide 'Gestão de Vulnerabilidades' '' 'Avaliação de Vulnerabilidades' (New-DeckKpiGrid $cards))
}

function New-DeckM365($Facts){
    $m=$Facts.manual
    $tb=$m.m365BackupStorageTB; $tbVal = if ($null -ne $tb){ "$tb TB" } else { '—' }
    $cards=@(
        @{n="$($m.m365SeatsUsed)/$($m.m365SeatsContracted)";cls='cyan';t='Assentos Protegidos';c='caixas de correio'},
        @{n="$($m.m365SeatsShared)";cls='blue';t='Compartilhados';c='assentos shared'},
        @{n="$($m.m365SharepointSites)";cls='gold';t='SharePoint';c='sites'},
        @{n="$tbVal";cls='orange';t='Backup M365';c='protegido'}
    )
    return (New-DeckSlide 'Nuvem Produtiva' '' 'Proteção Microsoft 365' (New-DeckKpiGrid $cards))
}

function New-DeckFrota($Facts){
    $lc=$Facts.lifecycle
    $onb = @($lc.onboardedThisMonth).Count
    $newest = $lc.newestVersion
    $newCount = (@($lc.versionBreakdown)|Where-Object{$_.version -eq $newest}|Select-Object -First 1).count
    $outdated = (@($lc.versionBreakdown)|Where-Object{$_.version -ne $newest}|Measure-Object count -Sum).Sum
    $mesNome = (_mesPt $Facts.meta.month).Split(' ')[0].ToLower()
    $box = "<div class='bluebox'><div class='h'>$($lc.online) de $($lc.total) agentes online</div><p>$($lc.total) agentes instalados$(if($onb){", com $onb novos em $mesNome"}). $([int]$lc.total - [int]$lc.online) seguem offline e precisam de atenção.</p></div>"
    $cards = "<div><div class='ocard'><div class='h'>$([int]$newCount) agentes</div><p>na versão mais recente $newest &#9989;</p></div><div class='ocard'><div class='h'>$([int]$outdated) agentes</div><p>em versões anteriores — atualização pendente</p></div></div>"
    $inner = "<div class='cols' style='margin-top:22px'>$box$cards</div>"
    return (New-DeckSlide 'Frota de Agentes' 'teal' 'Cobertura e Versões' $inner)
}

function New-DeckForaContrato($Facts){
    $esc=$script:DeckEsc
    $sb=New-Object System.Text.StringBuilder
    foreach($o in @($Facts.outOfContract)){
        [void]$sb.Append("<div class='ocard' style='border-color:#F0D030'><div class='h'>$(& $esc $o.labelPt)</div><p style='color:#F0D030'>$(& $esc $o.message)</p>")
        $m=$o.metrics
        if ($o.service -eq 'backup' -and $m){ [void]$sb.Append("<p>$($m.runs) execuções · $($m.error) falhas · $($m.successPct)% sucesso.</p>") }
        [void]$sb.Append("</div>")
    }
    return (New-DeckSlide 'Atividade Fora do Escopo' '' 'Pontos de Atenção — Fora do Contrato' $sb.ToString())
}

function New-DeckPlano($Facts){
    $esc=$script:DeckEsc
    $pr = @($Facts.analysis.topPriorities) | Select-Object -First 5
    $sb=New-Object System.Text.StringBuilder
    [void]$sb.Append("<ul class='alist' style='margin-top:18px'>")
    $i=0
    foreach($p in $pr){
        $i++
        $tag = if ("$p" -match 'offline|cr.tico|alerta|seguran|incidente'){'bad'}else{'warn'}
        $lab = if ("$p" -match 'patch|atualiz|reinicializ|disco|higiene'){'Higiene'} elseif("$p" -match 'armazen|cota|capacid'){'Capacidade'} elseif("$p" -match 'backup'){'Backup'} else {'Segurança'}
        [void]$sb.Append("<li><span class='num'>$i</span><span class='tx'><b>$(& $esc $p)</b></span><span class='tag $tag'>$lab</span></li>")
    }
    if (-not $pr){ [void]$sb.Append("<li><span class='tx'>Sem ações prioritárias no período.</span></li>") }
    [void]$sb.Append("</ul>")
    $inner = $sb.ToString() + "<div class='foot'><span>Contego Security · MSSP Acronis Cyber Protect Cloud</span><span>$(& $esc $Facts.meta.client) · $(_mesPt $Facts.meta.month)</span></div>"
    return (New-DeckSlide 'Próximos Passos' '' 'Plano de Ação' $inner)
}

function Build-DeckHtml {
    param([pscustomobject]$Facts,[string]$ConfigDir)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!doctype html><html lang='pt-BR'><head><meta charset='utf-8'>")
    [void]$sb.Append("<link rel='preconnect' href='https://fonts.googleapis.com'><link href='https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700;800&family=Inter:wght@400;500;600&display=swap' rel='stylesheet'>")
    [void]$sb.Append((Get-DeckCss))
    [void]$sb.Append("</head><body>")
    [void]$sb.Append((New-DeckCover $Facts $ConfigDir))
    [void]$sb.Append((New-DeckResumo $Facts))
    $vu = Get-DeckVisibleUsage $Facts
    if (@($vu).Count){ [void]$sb.Append((New-DeckServicos $vu)) }
    if ($Facts.serviceMap.backup -and [int]$Facts.backup.runs -gt 0){ [void]$sb.Append((New-DeckBackup $Facts)) }
    if ($Facts.serviceMap.edr -and $Facts.security){ [void]$sb.Append((New-DeckSeguranca $Facts)) }
    if ($Facts.serviceMap.va -and $Facts.vulnerability -and [int]$Facts.vulnerability.scansRun -gt 0){ [void]$sb.Append((New-DeckVa $Facts)) }
    if ($Facts.serviceMap.m365){ [void]$sb.Append((New-DeckM365 $Facts)) }
    if ($Facts.lifecycle -and [int]$Facts.lifecycle.total -gt 0){ [void]$sb.Append((New-DeckFrota $Facts)) }
    if (@($Facts.outOfContract).Count){ [void]$sb.Append((New-DeckForaContrato $Facts)) }
    [void]$sb.Append((New-DeckPlano $Facts))
    [void]$sb.Append("</body></html>")
    return $sb.ToString()
}

Export-ModuleMember -Function Build-DeckHtml, Get-DeckCss, New-DeckCover, Get-DeckLogo, New-DeckSlide, New-DeckKpiGrid, New-DeckCallout, New-DeckResumo, Get-DeckVisibleUsage, New-DeckServicos, New-DeckBackup, New-DeckSeguranca, New-DeckVa, New-DeckM365, New-DeckFrota, New-DeckForaContrato, New-DeckPlano