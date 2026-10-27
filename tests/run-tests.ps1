# Runs the DotNetVBA test suite: source lint, import-mode runs in Excel and/or Word,
# then the add-in boundary test (built .xlam referenced from a fresh workbook).
# Exit code = number of failures (2 when VBA project access is not trusted).
param(
    [ValidateSet('Excel', 'Word', 'All')][string]$Office = 'All',
    [switch]$SkipAddIn,
    [int]$TimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$srcDir = Join-Path $repoRoot 'src'
$testsDir = Join-Path $repoRoot 'tests'

# Finds modal dialogs (e.g. a VBA compile error) owned by an Office process,
# so a broken build fails fast instead of blocking Application.Run.
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class DotNetVbaDialogWatch {
    private delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, StringBuilder lParam);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    private const uint WM_GETTEXT = 0x000D;
    private const uint WM_CLOSE = 0x0010;

    private static string TextOf(IntPtr hWnd) {
        var sb = new StringBuilder(2048);
        SendMessage(hWnd, WM_GETTEXT, (IntPtr)sb.Capacity, sb);
        return sb.ToString();
    }

    // Closes every visible dialog of the process; returns their texts.
    public static string[] CloseDialogs(uint processId) {
        var found = new List<string>();
        EnumWindows((hWnd, l) => {
            uint pid;
            GetWindowThreadProcessId(hWnd, out pid);
            if (pid != processId || !IsWindowVisible(hWnd)) return true;
            var cls = new StringBuilder(256);
            GetClassName(hWnd, cls, cls.Capacity);
            if (cls.ToString() != "#32770") return true;
            var parts = new List<string>();
            parts.Add(TextOf(hWnd));
            EnumChildWindows(hWnd, (child, l2) => {
                var childCls = new StringBuilder(256);
                GetClassName(child, childCls, childCls.Capacity);
                if (childCls.ToString() == "Static") {
                    string t = TextOf(child);
                    if (t.Length > 0) parts.Add(t);
                }
                return true;
            }, IntPtr.Zero);
            found.Add(string.Join(" | ", parts).Replace("\r", " ").Replace("\n", " "));
            PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }
}
"@

function Invoke-Lint {
    $files = @(Get-ChildItem -Path (Join-Path $srcDir '*.cls'), (Join-Path $srcDir '*.bas'), (Join-Path $testsDir '*.bas') -File -ErrorAction SilentlyContinue)
    $problems = 0
    foreach ($file in $files) {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            Write-Host "LINT FAIL $($file.Name): starts with a UTF-8 BOM"
            $problems++
        }
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -ge 0x80) {
                Write-Host "LINT FAIL $($file.Name): non-ASCII byte 0x$('{0:X2}' -f $bytes[$i]) at offset $i"
                $problems++
                break
            }
        }
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -eq 0x0A -and ($i -eq 0 -or $bytes[$i - 1] -ne 0x0D)) {
                Write-Host "LINT FAIL $($file.Name): LF without CR at offset $i"
                $problems++
                break
            }
        }
    }
    if ($problems -eq 0) { Write-Host "LINT OK ($($files.Count) files)" }
    return $problems
}

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# Background thread: if the Office process shows a dialog while tests run, or the
# run exceeds the timeout, records why and kills the process.
function Start-Watchdog([int]$ProcessId, $State) {
    Start-ThreadJob -ArgumentList $ProcessId, $State, $TimeoutSeconds -ScriptBlock {
        param($ProcessId, $State, $TimeoutSeconds)
        $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
        while (-not $State.Done) {
            $texts = [DotNetVbaDialogWatch]::CloseDialogs([uint32]$ProcessId)
            if ($texts.Count -gt 0) {
                # A dialog means the run is broken (closing a compile error
                # leaves VBA in break mode with Run still blocked), so stop it.
                foreach ($text in $texts) { [void]$State.Dialogs.Add($text) }
                Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
                break
            }
            if ((Get-Date) -gt $deadline) {
                [void]$State.Dialogs.Add("timed out after $TimeoutSeconds s; process killed")
                Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
                break
            }
            Start-Sleep -Milliseconds 300
        }
    }
}

# Starts an invisible Office application with alerts off and the dialog
# watchdog attached. Always pair with Stop-OfficeSession in a finally block.
function Start-OfficeSession {
    param([ValidateSet('Excel', 'Word')][string]$App)

    $session = @{
        App       = $App
        OfficeApp = $null
        Docs      = $null
        Watchdog  = $null
        NewPids   = @()
        State     = [hashtable]::Synchronized(@{
            Done    = $false
            Dialogs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        })
    }
    $processName = if ($App -eq 'Excel') { 'EXCEL' } else { 'WINWORD' }
    $before = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object Id)
    $session.OfficeApp = New-Object -ComObject "$App.Application"
    $session.NewPids = @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
                         Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
    $session.OfficeApp.Visible = $false
    if ($App -eq 'Excel') {
        $session.OfficeApp.DisplayAlerts = $false
        $session.Docs = $session.OfficeApp.Workbooks
    } else {
        $session.OfficeApp.DisplayAlerts = 0
        $session.Docs = $session.OfficeApp.Documents
    }
    if ($session.NewPids.Count -eq 1) {
        $session.Watchdog = Start-Watchdog -ProcessId $session.NewPids[0] -State $session.State
    } else {
        Write-Host "WARN: could not identify the $App process; dialog watchdog disabled"
    }
    return $session
}

# Closes a document or workbook without saving and releases it; ignores a dead process.
function Close-OfficeDocument($Session, $Doc) {
    if ($null -eq $Doc) { return }
    try {
        if ($Session.App -eq 'Excel') { $Doc.Close($false) } else { $Doc.Close(0) }
    } catch { }
    Release-ComObject $Doc
}

# Quits the application, releases COM, stops the watchdog and kills the process
# if it does not exit. Callers close and release their own objects first.
function Stop-OfficeSession($Session) {
    if ($null -eq $Session) { return }
    Release-ComObject $Session.Docs
    $Session.Docs = $null
    if ($null -ne $Session.OfficeApp) {
        try { $Session.OfficeApp.Quit() } catch { }
        Release-ComObject $Session.OfficeApp
        $Session.OfficeApp = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $Session.State.Done = $true
    if ($null -ne $Session.Watchdog) {
        Wait-Job $Session.Watchdog -Timeout 10 | Out-Null
        Remove-Job $Session.Watchdog -Force
    }
    foreach ($leftover in $Session.NewPids) {
        $proc = Get-Process -Id $leftover -ErrorAction SilentlyContinue
        if ($null -ne $proc -and -not $proc.WaitForExit(10000)) {
            Write-Host "WARN: $($Session.App) process $leftover did not exit; killing it"
            Stop-Process -Id $leftover -Force -ErrorAction SilentlyContinue
        }
    }
}

# Imports a module file; returns $null on success or the failure text.
function Import-VbaFile($Components, $File) {
    try {
        $component = $Components.Import($File.FullName)
        $importedName = $component.Name
        Release-ComObject $component
    } catch {
        return "import of $($File.Name) failed: $($_.Exception.Message)"
    }
    if ($importedName -ne $File.BaseName) {
        return "$($File.Name) imported as '$importedName' (check its Attribute VB_Name header)"
    }
    return $null
}

# Runs a test macro (trying each candidate name in turn) that writes a UTF-16
# log, prints any dialog the watchdog closed and the log, and returns the
# failure count (at least 1 when the macro could not run or a dialog appeared).
function Invoke-LoggedMacro($Session, [string[]]$MacroNames, [string]$Log, [string]$Label) {
    if (Test-Path $Log) { Remove-Item $Log }
    $failures = $null
    $lastError = $null
    foreach ($macro in $MacroNames) {
        try {
            $failures = [int]$Session.OfficeApp.Run($macro, $Log)
            break
        } catch {
            $lastError = $_.Exception.Message
            # The watchdog killed the process; no other name can run.
            if ($Session.State.Dialogs.Count -gt 0) { break }
        }
    }
    if ($null -eq $failures) {
        Write-Host "$Label FAIL: $($MacroNames[0]) could not run: $lastError"
        $failures = 1
    }
    foreach ($text in $Session.State.Dialogs) {
        Write-Host "$Label FAIL: dialog shown during run: $text"
    }
    if ($Session.State.Dialogs.Count -gt 0 -and $failures -eq 0) { $failures = 1 }
    if (Test-Path $Log) { Get-Content $Log -Encoding Unicode | ForEach-Object { Write-Host $_ } }
    return $failures
}

function Invoke-ImportModeRun {
    param([ValidateSet('Excel', 'Word')][string]$App)

    Write-Host "=== $App (import mode) ==="
    $session = $null; $doc = $null; $proj = $null; $components = $null
    try {
        $session = Start-OfficeSession -App $App
        $doc = $session.Docs.Add()

        try {
            $proj = $doc.VBProject
            $components = $proj.VBComponents
        } catch {
            Write-Host "Enable `"Trust access to the VBA project object model`" in $App Trust Center"
            return -1
        }

        $paths = @(Get-ChildItem -Path (Join-Path $srcDir '*.bas') -File) +
                 @(Get-ChildItem -Path (Join-Path $srcDir '*.cls') -File) +
                 @(Get-Item (Join-Path $testsDir 'TestDotNetVBA.bas'))
        foreach ($p in $paths) {
            $problem = Import-VbaFile $components $p
            if ($null -ne $problem) {
                Write-Host "$App FAIL: $problem"
                return 1
            }
        }

        $log = Join-Path $env:TEMP "dotnetvba-$App.log"
        return Invoke-LoggedMacro $session @('RunAllTests') $log $App
    } finally {
        Release-ComObject $components
        Release-ComObject $proj
        if ($null -ne $session) { Close-OfficeDocument $session $doc }
        $components = $null; $proj = $null; $doc = $null
        Stop-OfficeSession $session
    }
}

# Builds the add-in to a temp folder, references it from a fresh workbook and
# runs TestAddInConsumer there, so the public API is exercised across the
# project boundary. Returns the failure count.
function Invoke-AddInRun {
    Write-Host '=== Excel (add-in) ==='
    $testDir = Join-Path $env:TEMP 'dotnetvba-test'
    $xlam = Join-Path $testDir 'DotNetVBA.xlam'
    $buildScript = Join-Path $repoRoot 'build\build-xlam.ps1'
    try {
        & $buildScript -OutputPath $xlam
        $buildExit = $LASTEXITCODE
    } catch {
        Write-Host "ADDIN FAIL: build-xlam.ps1 threw: $($_.Exception.Message)"
        return 1
    }
    if ($buildExit -ne 0 -or -not (Test-Path $xlam)) {
        Write-Host "ADDIN FAIL: build-xlam.ps1 failed (exit $buildExit)"
        return 1
    }

    $session = $null; $addIn = $null; $wb = $null; $proj = $null; $components = $null; $refs = $null; $ref = $null
    try {
        $session = Start-OfficeSession -App 'Excel'
        try {
            $addIn = $session.Docs.Open($xlam)
            $wb = $session.Docs.Add()
            $proj = $wb.VBProject
            $components = $proj.VBComponents
            $refs = $proj.References
            $ref = $refs.AddFromFile($xlam)
        } catch {
            Write-Host "ADDIN FAIL: could not open or reference the add-in: $($_.Exception.Message)"
            foreach ($text in $session.State.Dialogs) { Write-Host "ADDIN FAIL: dialog shown: $text" }
            return 1
        }

        $problem = Import-VbaFile $components (Get-Item (Join-Path $testsDir 'TestAddInConsumer.bas'))
        if ($null -ne $problem) {
            Write-Host "ADDIN FAIL: $problem"
            return 1
        }

        $log = Join-Path $testDir 'dotnetvba-addin.log'
        $qualified = "'" + $wb.Name + "'!RunAddInConsumerTest"
        return Invoke-LoggedMacro $session @('RunAddInConsumerTest', $qualified) $log 'ADDIN'
    } finally {
        Release-ComObject $ref
        Release-ComObject $refs
        Release-ComObject $components
        Release-ComObject $proj
        $ref = $null; $refs = $null; $components = $null; $proj = $null
        if ($null -ne $session) {
            Close-OfficeDocument $session $wb
            Close-OfficeDocument $session $addIn
        }
        $wb = $null; $addIn = $null
        Stop-OfficeSession $session
    }
}

$total = Invoke-Lint

$hosts = if ($Office -eq 'All') { @('Excel', 'Word') } else { @($Office) }
foreach ($h in $hosts) {
    $result = Invoke-ImportModeRun -App $h
    if ($result -lt 0) { exit 2 }
    $total += $result
}

if (-not $SkipAddIn -and $Office -ne 'Word') {
    $addInFailures = Invoke-AddInRun
    Write-Host "ADDIN: $addInFailures failures"
    $total += $addInFailures
}

Write-Host "EXIT $total"
exit $total
