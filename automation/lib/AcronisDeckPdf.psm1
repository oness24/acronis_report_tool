# automation/lib/AcronisDeckPdf.psm1 — render the deck HTML to a landscape PDF with a headless Chromium/Edge.
$ErrorActionPreference = "Stop"

function Resolve-DeckBrowser {
    # PATH-resolvable Chromium-family commands, then known Windows Edge/Chrome locations.
    foreach($n in 'chromium','chromium-browser','google-chrome','chrome','msedge'){
        $c = Get-Command $n -ErrorAction SilentlyContinue
        if ($c){ return $c.Source }
    }
    $cands = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    foreach($p in $cands){ if ($p -and (Test-Path $p)){ return $p } }
    return $null
}

function Convert-DeckToPdf {
    param([string]$HtmlPath,[string]$PdfPath)
    $browser = Resolve-DeckBrowser
    if (-not $browser){ Write-Warning "  deck: navegador (Chromium/Edge) ausente - pulando PDF do deck"; return $false }
    $errLog = $null; $outLog = $null
    try {
        # URI/temp-file setup can throw (UriFormatException on odd paths; IOException if temp is full) - keep inside try.
        $uri = ([Uri]("file://" + ($HtmlPath -replace '\\','/'))).AbsoluteUri
        $errLog = [System.IO.Path]::GetTempFileName(); $outLog = [System.IO.Path]::GetTempFileName()
        $proc = Start-Process -FilePath $browser `
            -ArgumentList "--headless","--disable-gpu","--disable-logging","--log-level=3","--no-sandbox","--no-pdf-header-footer","--print-to-pdf=$PdfPath",$uri `
            -PassThru -NoNewWindow -RedirectStandardError $errLog -RedirectStandardOutput $outLog
        $proc.WaitForExit(45000) | Out-Null
        if (-not $proc.HasExited){ try { $proc.Kill() } catch {} }
    } catch {
        Write-Warning "  deck: falha ao gerar PDF: $($_.Exception.Message)"; return $false
    } finally {
        # null-safe: Remove-Item binds null/empty -Path with a TERMINATING error that
        # -ErrorAction can't suppress, so skip nulls (set when GetTempFileName threw).
        foreach($l in @($errLog,$outLog)){ if ($l){ Remove-Item $l -ErrorAction SilentlyContinue } }
    }
    return [bool](Test-Path $PdfPath)
}

Export-ModuleMember -Function Resolve-DeckBrowser, Convert-DeckToPdf
