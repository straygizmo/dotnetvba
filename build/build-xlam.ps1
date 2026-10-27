# Builds DotNetVBA.xlam from src\: imports every module and class into a new
# workbook and saves it as an Excel add-in.
# Prints "BUILT <path> (DotNetVBA <version>)" and exits 0; exits 1 on any failure.
param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $repoRoot 'src'
if ([string]::IsNullOrEmpty($OutputPath)) {
    $OutputPath = Join-Path $repoRoot 'dist\DotNetVBA.xlam'
}
$XL_OPEN_XML_ADDIN = 55
# Shown in Tools > References. VBA forbids a module with the project's name,
# which is why the facade class is DotNet rather than DotNetVBA.
$PROJECT_NAME = 'DotNetVBA'

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# DocumentProperties is late-bound only, so PowerShell's adapter cannot call
# Item on it; go through IDispatch by reflection instead.
function Set-DocumentProperty($properties, [string]$name, [string]$value) {
    $prop = [System.__ComObject].InvokeMember('Item', 'GetProperty', $null, $properties, @($name))
    try {
        [void][System.__ComObject].InvokeMember('Value', 'SetProperty', $null, $prop, @($value))
    } finally {
        Release-ComObject $prop
    }
}

$excel = $null; $books = $null; $wb = $null; $proj = $null; $components = $null; $props = $null
$newPids = @()
$exitCode = 1
try {
    $match = Select-String -Path (Join-Path $srcDir 'DotNet.cls') -Pattern 'VERSION_STRING As String = "([^"]+)"'
    if ($null -eq $match) { throw 'VERSION_STRING not found in src\DotNet.cls' }
    $version = $match.Matches[0].Groups[1].Value

    $fullPath = [System.IO.Path]::GetFullPath($OutputPath)
    $outDir = Split-Path -Parent $fullPath
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (Test-Path $fullPath) { Remove-Item $fullPath -Force }

    $before = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object Id)
    $excel = New-Object -ComObject 'Excel.Application'
    $newPids = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue |
                 Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $books = $excel.Workbooks
    $wb = $books.Add()
    try {
        $proj = $wb.VBProject
        $components = $proj.VBComponents
    } catch {
        throw 'cannot access the VBA project; enable "Trust access to the VBA project object model" in Excel Trust Center'
    }
    $paths = @(Get-ChildItem -Path (Join-Path $srcDir '*.bas') -File) +
             @(Get-ChildItem -Path (Join-Path $srcDir '*.cls') -File)
    foreach ($p in $paths) {
        $component = $components.Import($p.FullName)
        $importedName = $component.Name
        Release-ComObject $component
        if ($importedName -ne $p.BaseName) {
            throw "$($p.Name) imported as '$importedName' (check its Attribute VB_Name header)"
        }
    }

    $proj.Name = $PROJECT_NAME

    $props = $wb.BuiltinDocumentProperties
    Set-DocumentProperty $props 'Title' 'DotNetVBA'
    Set-DocumentProperty $props 'Comments' "DotNetVBA $version - .NET-style API for VBA"

    $wb.SaveAs($fullPath, $XL_OPEN_XML_ADDIN)
    if (-not (Test-Path $fullPath)) { throw "SaveAs did not create $fullPath" }
    Write-Host "BUILT $fullPath (DotNetVBA $version)"
    $exitCode = 0
} catch {
    Write-Host "BUILD FAIL: $($_.Exception.Message) [line $($_.InvocationInfo.ScriptLineNumber)]"
    $exitCode = 1
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    Release-ComObject $props
    Release-ComObject $components
    Release-ComObject $proj
    Release-ComObject $wb
    Release-ComObject $books
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        Release-ComObject $excel
    }
    $props = $null; $components = $null; $proj = $null; $wb = $null; $books = $null; $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    foreach ($leftover in $newPids) {
        $proc = Get-Process -Id $leftover -ErrorAction SilentlyContinue
        if ($null -ne $proc -and -not $proc.WaitForExit(10000)) {
            Write-Host "WARN: Excel process $leftover did not exit; killing it"
            Stop-Process -Id $leftover -Force -ErrorAction SilentlyContinue
        }
    }
}
exit $exitCode
