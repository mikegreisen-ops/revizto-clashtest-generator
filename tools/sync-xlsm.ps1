<#
.SYNOPSIS
    Inject the src\*.bas modules into a macro-enabled workbook, so the workbook's
    embedded macros always match the canonical .bas source.

.DESCRIPTION
    The .bas files in src\ are the single source of truth for the macros. A workbook
    (.xlsm) holds its OWN embedded copy of each module, which does NOT update when a
    .bas changes. This script re-imports every src\*.bas into the target workbook by
    replacing the module of the same name (clean overwrite, by module name).

    Default target is templates\Example.xlsm (the shipped starter). Point -Workbook at
    a scratch test workbook to refresh that instead, e.g.:
        powershell -File tools\sync-xlsm.ps1 -Workbook scratch\test.xlsm

.NOTES
    Requires Excel + the one-time setting:
        Excel > File > Options > Trust Center > Trust Center Settings >
        Macro Settings > [x] Trust access to the VBA project object model
    The target workbook must be CLOSED in Excel when this runs.
#>

[CmdletBinding()]
param(
    [string]$Workbook,
    [string]$SrcDir
)

$ErrorActionPreference = 'Stop'
$vbext_ct_StdModule = 1

# repo root = parent of this tools\ folder
$repoRoot = Split-Path $PSScriptRoot -Parent
if (-not $Workbook) { $Workbook = Join-Path $repoRoot 'templates\Example.xlsm' }
if (-not $SrcDir)   { $SrcDir   = Join-Path $repoRoot 'src' }

# resolve to an absolute path (Excel COM rejects relative paths).
# a relative -Workbook is resolved against the repo root.
if (-not [System.IO.Path]::IsPathRooted($Workbook)) {
    $Workbook = Join-Path $repoRoot $Workbook
}
$wbPath = [System.IO.Path]::GetFullPath($Workbook)
if (-not (Test-Path -LiteralPath $wbPath)) { throw "Workbook not found: $wbPath" }

$basFiles = @(Get-ChildItem -LiteralPath $SrcDir -Filter *.bas -ErrorAction SilentlyContinue)
if ($basFiles.Count -eq 0) { throw "No .bas files found in $SrcDir" }

Write-Host "Target workbook : $wbPath"
Write-Host "Source modules  : $($basFiles.Name -join ', ')"
Write-Host ""

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Open($wbPath)

    # verify programmatic VBA access is trusted
    $proj = $null
    try { $proj = $wb.VBProject } catch { }
    if ($null -eq $proj) {
        throw @"
Cannot access the workbook's VBA project.
Enable it once:  Excel > File > Options > Trust Center > Trust Center Settings >
                 Macro Settings > [x] Trust access to the VBA project object model
"@
    }

    foreach ($bas in $basFiles) {
        $modName = [System.IO.Path]::GetFileNameWithoutExtension($bas.Name)

        # remove any existing module of that name
        $toRemove = @()
        foreach ($comp in $proj.VBComponents) {
            if ($comp.Name -eq $modName) { $toRemove += $comp }
        }
        foreach ($comp in $toRemove) { $proj.VBComponents.Remove($comp) }

        # add a fresh standard module and load the .bas text
        $comp = $proj.VBComponents.Add($vbext_ct_StdModule)
        $comp.Name = $modName
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }  # drop any auto "Option Explicit"
        $code = Get-Content -Raw -LiteralPath $bas.FullName
        $cm.AddFromString($code)

        Write-Host ("  synced  {0}  ({1} lines)" -f $modName, $cm.CountOfLines)
    }

    # warn about any standard module NOT backed by a src\*.bas (e.g. a stale "Module1"
    # left over from an original paste) - these can cause "Ambiguous name detected".
    $managed = $basFiles | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension($_.Name) }
    $strays = @()
    foreach ($comp in $proj.VBComponents) {
        if ($comp.Type -eq $vbext_ct_StdModule -and ($managed -notcontains $comp.Name)) {
            $strays += $comp.Name
        }
    }
    if ($strays.Count -gt 0) {
        Write-Warning ("Unmanaged module(s) present (not from src\): {0}. " -f ($strays -join ', ')) `
            + "If these are old copies of the macros, delete them in the VBA editor to avoid name clashes."
    }

    $wb.Save()
    Write-Host ""
    Write-Host "Done. $($basFiles.Count) module(s) written to $([System.IO.Path]::GetFileName($wbPath))."
}
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($xl) { try { $xl.Quit() } catch { } }
    foreach ($o in @($wb, $xl)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
