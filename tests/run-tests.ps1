# Runs the DotNetVBA test suite: source lint, then import-mode runs in Excel and/or Word.
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

function Invoke-ImportModeRun {
    param([ValidateSet('Excel', 'Word')][string]$App)

    Write-Host "=== $App (import mode) ==="
    $processName = if ($App -eq 'Excel') { 'EXCEL' } else { 'WINWORD' }
    $officeApp = $null; $docs = $null; $doc = $null; $proj = $null; $components = $null
    $watchdog = $null
    $newPids = @()
    $state = [hashtable]::Synchronized(@{
        Done    = $false
        Dialogs = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    })
    $failures = 0
    try {
        $before = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object Id)
        $officeApp = New-Object -ComObject "$App.Application"
        $officeApp.Visible = $false
        if ($App -eq 'Excel') {
            $officeApp.DisplayAlerts = $false
            $docs = $officeApp.Workbooks
        } else {
            $officeApp.DisplayAlerts = 0
            $docs = $officeApp.Documents
        }
        $newPids = @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
                     Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
        if ($newPids.Count -eq 1) {
            $watchdog = Start-Watchdog -ProcessId $newPids[0] -State $state
        } else {
            Write-Host "WARN: could not identify the $App process; dialog watchdog disabled"
        }
        $doc = $docs.Add()

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
            try {
                $component = $components.Import($p.FullName)
                $importedName = $component.Name
                Release-ComObject $component
            } catch {
                Write-Host "$App FAIL: import of $($p.Name) failed: $($_.Exception.Message)"
                return 1
            }
            if ($importedName -ne $p.BaseName) {
                Write-Host "$App FAIL: $($p.Name) imported as '$importedName' (check its Attribute VB_Name header)"
                return 1
            }
        }

        $log = Join-Path $env:TEMP "dotnetvba-$App.log"
        if (Test-Path $log) { Remove-Item $log }
        try {
            $failures = [int]$officeApp.Run('RunAllTests', $log)
        } catch {
            Write-Host "$App FAIL: RunAllTests could not run: $($_.Exception.Message)"
            $failures = 1
        }
        foreach ($text in $state.Dialogs) {
            Write-Host "$App FAIL: dialog shown during run: $text"
        }
        if ($state.Dialogs.Count -gt 0 -and $failures -eq 0) { $failures = 1 }
        # VBA's TextStream writes the ANSI code page.
        if (Test-Path $log) { Get-Content $log -Encoding ansi | ForEach-Object { Write-Host $_ } }
    } finally {
        if ($null -ne $doc) {
            try {
                if ($App -eq 'Excel') { $doc.Close($false) } else { $doc.Close(0) }
            } catch { }
        }
        Release-ComObject $components
        Release-ComObject $proj
        Release-ComObject $doc
        Release-ComObject $docs
        if ($null -ne $officeApp) {
            try { $officeApp.Quit() } catch { }
            Release-ComObject $officeApp
        }
        $components = $null; $proj = $null; $doc = $null; $docs = $null; $officeApp = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        $state.Done = $true
        if ($null -ne $watchdog) {
            Wait-Job $watchdog -Timeout 10 | Out-Null
            Remove-Job $watchdog -Force
        }
        foreach ($leftover in $newPids) {
            $proc = Get-Process -Id $leftover -ErrorAction SilentlyContinue
            if ($null -ne $proc -and -not $proc.WaitForExit(10000)) {
                Write-Host "WARN: $App process $leftover did not exit; killing it"
                Stop-Process -Id $leftover -Force -ErrorAction SilentlyContinue
            }
        }
    }
    return $failures
}

function Invoke-AddInRun {
    Write-Host 'ADDIN: skipped (not built yet)'
    return 0
}

$total = Invoke-Lint

$hosts = if ($Office -eq 'All') { @('Excel', 'Word') } else { @($Office) }
foreach ($h in $hosts) {
    $result = Invoke-ImportModeRun -App $h
    if ($result -lt 0) { exit 2 }
    $total += $result
}

if (-not $SkipAddIn -and $Office -ne 'Word') {
    $total += Invoke-AddInRun
}

Write-Host "EXIT $total"
exit $total
