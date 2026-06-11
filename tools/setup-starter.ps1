<#
.SYNOPSIS
    Build the starter workbook (templates/Example.xlsm) from scratch, reproducibly:
    imports the src\*.bas modules, and (re)creates the Instructions, Sets and Tests sheets
    plus the buttons on the Instructions sheet. Run this instead of hand-editing the .xlsm.

.NOTES
    Requires Excel + the one-time "Trust access to the VBA project object model" setting
    (Excel > Options > Trust Center > Macro Settings). The workbook must be closed.
    The Tests-sheet layout here mirrors LayOutTestsSheet in modClashLite.bas - keep them in step.
#>
[CmdletBinding()]
param([string]$Workbook)

$ErrorActionPreference = 'Stop'
$xlCenter = -4108
$vbext_ct_StdModule = 1
$GREY  = 4210752       # RGB(64,64,64)
$WHITE = 16777215

$repoRoot = Split-Path $PSScriptRoot -Parent
if (-not $Workbook) { $Workbook = Join-Path $repoRoot 'templates\Example.xlsm' }
$wbPath = [System.IO.Path]::GetFullPath($Workbook)
$srcDir = Join-Path $repoRoot 'src'
$basFiles = @(Get-ChildItem -LiteralPath $srcDir -Filter *.bas)
if ($basFiles.Count -eq 0) { throw "No .bas files in $srcDir" }

function Get-Sheet($wb, $name) {
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq $name) { return $s } }
    $s = $wb.Worksheets.Add()
    $s.Name = $name
    return $s
}

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false

    if (Test-Path -LiteralPath $wbPath) { $wb = $xl.Workbooks.Open($wbPath) }
    else { $wb = $xl.Workbooks.Add(); $wb.SaveAs($wbPath, 52) }   # 52 = xlOpenXMLWorkbookMacroEnabled

    # ---- 1. import the VBA modules (replace same-named) ----
    $proj = $null
    try { $proj = $wb.VBProject } catch {}
    if ($null -eq $proj) {
        throw "Can't access the VBA project. Enable: Excel > Trust Center > Macro Settings > Trust access to the VBA project object model."
    }
    foreach ($bas in $basFiles) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($bas.Name)
        $rm = @(); foreach ($c in $proj.VBComponents) { if ($c.Name -eq $name) { $rm += $c } }
        foreach ($c in $rm) { $proj.VBComponents.Remove($c) }
        $comp = $proj.VBComponents.Add($vbext_ct_StdModule)
        $comp.Name = $name
        $comp.CodeModule.AddFromString((Get-Content -Raw -LiteralPath $bas.FullName))
        Write-Host "  module  $name"
    }

    # ---- 2. Sets sheet ----
    $sets = Get-Sheet $wb 'Sets'
    $sets.Cells.Clear() | Out-Null
    $sets.Range('A1').Value2 = 'Search Set Names'
    $sets.Range('A1').Font.Bold = $true
    $sets.Columns.Item('A').ColumnWidth = 38

    # ---- 3. Tests sheet (mirrors LayOutTestsSheet) ----
    $t = Get-Sheet $wb 'Tests'
    $t.Cells.Clear() | Out-Null
    $hdr = 'Test Name','Tolerance (mm)','Clearance (mm)','Priority','Stamp','Type'
    for ($i = 0; $i -lt $hdr.Count; $i++) { $t.Cells.Item(1, $i + 1).Value2 = $hdr[$i] }
    $t.Cells.Item(1,25).Value2 = 'Set A'      # Y
    $t.Cells.Item(1,26).Value2 = 'Set B'      # Z
    $t.Columns.Item('A').AutoFit() | Out-Null
    $t.Columns.Item('B:F').ColumnWidth = 24
    $t.Columns.Item('B:F').HorizontalAlignment = $xlCenter
    $th = $t.Range('A1:F1')
    $th.HorizontalAlignment = $xlCenter
    $th.Font.Bold = $true; $th.Font.Color = $WHITE; $th.Interior.Color = $GREY
    $t.Range($t.Cells.Item(1,25), $t.Cells.Item(1,26)).EntireColumn.Hidden = $true
    if ($t.AutoFilterMode) { $t.AutoFilterMode = $false }
    $t.Range('A1').AutoFilter() | Out-Null      # header-row filter dropdowns (incl. Type)

    # ---- 4. Instructions sheet (text + buttons) ----
    $ins = Get-Sheet $wb 'Instructions'
    $ins.Cells.Clear() | Out-Null
    if ($ins.Buttons().Count -gt 0) { $ins.Buttons().Delete() | Out-Null }
    $lines = @(
        @('Revizto Clash-Test Generator', 18, $true),
        @('Build a Revizto clash test for every pairing of your search sets - set their tolerance, clearance, priority and stamp - then export a .vimctst to import into Revizto.', 11, $false),
        @('', 11, $false),
        @('Getting started', 12, $true),
        @('1.  Put your search-set names on the Sets sheet (column A). No list yet? Click "Import Set Names" and pick a .vimsst export from Revizto.', 11, $false),
        @('2.  Click "Build Test List" to create every pairing - or paste your own list onto the Tests sheet (it is ready from the start).', 11, $false),
        @('3.  On the Tests sheet, set Tolerance / Clearance / Priority / Stamp. Use the Type column to filter and delete Self-pairs if you do not want them.', 11, $false),
        @('4.  Click "Export Clash Tests", choose where to save, and import the .vimctst into Revizto (Clash Detection > import).', 11, $false),
        @('', 11, $false),
        @('Editing tests you already have: click "Import Clash Tests", pick the file, edit, then "Export Clash Tests".', 11, $false),
        @('', 11, $false),
        @('Good to know', 12, $true),
        @('-  Set names and stamp codes must match your Revizto project exactly. Stamp codes are CASE-SENSITIVE (0AR is not 0ar).', 11, $false),
        @('-  Sets not in Revizto yet? Use Revizto''s Bulk Create Search Sets to make them first.', 11, $false),
        @('-  Re-importing makes NEW tests - delete the originals in Revizto first, or you will get duplicates.', 11, $false),
        @('-  Tolerance and Clearance are opposites - a test uses one or the other, never both.', 11, $false),
        @('-  Unofficial community tool: keep backups. Revizto have approved sharing it but do not support it.', 11, $false)
    )
    $r = 1
    foreach ($ln in $lines) {
        $cell = $ins.Cells.Item($r, 1)
        $cell.Value2 = $ln[0]; $cell.Font.Size = $ln[1]; $cell.Font.Bold = $ln[2]
        $r++
    }
    $ins.Columns.Item('A').ColumnWidth = 95
    $xl.ActiveWindow.DisplayGridlines = $false

    # buttons stacked to the right of the text
    $btnLeft = $ins.Range('I2').Left
    $btns = @(
        @('Import Set Names',  'ImportSetsFromVimsst'),
        @('Build Test List',   'BuildTestList'),
        @('New Blank Tests Sheet', 'NewTestsSheet'),
        @('Export Clash Tests','ExportClashTests'),
        @('Import Clash Tests','ImportClashTests')
    )
    $top = $ins.Range('A2').Top
    foreach ($b in $btns) {
        $btn = $ins.Buttons().Add($btnLeft, $top, 190, 32)
        $btn.Caption = $b[0]; $btn.OnAction = $b[1]; $btn.Font.Size = 11
        $top += 42
    }

    # ---- 5. drop any stray sheets, then order: Instructions, Sets, Tests; open on Instructions ----
    $keep = 'Instructions', 'Sets', 'Tests'
    $stray = @(); foreach ($s in $wb.Worksheets) { if ($keep -notcontains $s.Name) { $stray += $s } }
    foreach ($s in $stray) { $s.Delete() | Out-Null }
    $ins.Move($wb.Worksheets.Item(1))           # Instructions first
    $sets.Move([System.Reflection.Missing]::Value, $ins)
    $t.Move([System.Reflection.Missing]::Value, $sets)
    $ins.Activate()
    $ins.Range('A1').Select() | Out-Null

    $wb.Save()
    Write-Host ("Done. Sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ', '))
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
