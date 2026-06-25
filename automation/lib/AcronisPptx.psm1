# automation/lib/AcronisPptx.psm1
# Graceful PPTX integration -- Python/python-pptx is OPTIONAL.
# If absent, Invoke-PptxBuild warns once and returns; the pipeline continues normally.

function Resolve-PythonLauncher {
    <#
    .SYNOPSIS
      Returns the first Python launcher (as a string array) that can import python-pptx,
      or $null if none is found.  Never throws.
    .NOTES
      Tries: python, py -3, python3 (cross-platform: Windows + Linux/Colab).
      PS 5.1 note: $c[1..($c.Count-1)] when Count=1 yields $c[1..0] = reversed slice
      (a PS 5.1 quirk).  We guard with an explicit Count check to avoid passing garbage args.
    #>
    $candidates = @(
        @('python'),
        @('py', '-3'),
        @('python3')
    )
    foreach ($c in $candidates) {
        try {
            # PS 5.1: $c[1..0] reverses -- guard with explicit Count check
            $extraArgs = if ($c.Count -gt 1) { $c[1..($c.Count-1)] } else { @() }
            & $c[0] @extraArgs -c "import pptx" 2>$null
            if ($LASTEXITCODE -eq 0) { return ,$c }
        } catch {}
    }
    return $null
}

function Invoke-PptxBuild {
    <#
    .SYNOPSIS
      Calls build_pptx.py to produce a .pptx deck.
      Silently skips (warn only) when $Launcher is $null -- never throws.
    .PARAMETER FactsPath   Path to the facts_<slug>.json for this client.
    .PARAMETER OutDir      Directory where the .pptx is written.
    .PARAMETER ConfigDir   Path to the config/ directory (for branding assets).
    .PARAMETER Launcher    Launcher array from Resolve-PythonLauncher, or $null to skip.
    #>
    param(
        [string]$FactsPath,
        [string]$OutDir,
        [string]$ConfigDir,
        [object]$Launcher
    )
    if (-not $Launcher) {
        Write-Warning "  pptx: Python/python-pptx ausente - pulando deck"
        return
    }
    # build_pptx.py lives in automation/ which is one level above automation/lib/ ($PSScriptRoot)
    $script = Join-Path $PSScriptRoot "../build_pptx.py"
    $script = [System.IO.Path]::GetFullPath($script)
    # PS 5.1: guard against reversed-slice when launcher has no extra args (e.g. 'python')
    $launcherExtra = if ($Launcher.Count -gt 1) { $Launcher[1..($Launcher.Count-1)] } else { @() }
    $callArgs = @($launcherExtra) + @($script, $FactsPath, $OutDir, $ConfigDir)
    try {
        & $Launcher[0] @callArgs
    } catch {
        Write-Warning "  pptx: falha ao gerar deck: $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function Resolve-PythonLauncher, Invoke-PptxBuild
