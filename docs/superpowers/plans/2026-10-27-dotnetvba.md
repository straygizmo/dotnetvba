# DotNetVBA Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build DotNetVBA, a VBA class library with .NET API names (`Net.File`, `Net.Directory`, `Net.Path`, `Net.Environment`, `Net.List`, …), tested automatically against real Excel and Word, packaged as an `.xlam`, ready to tag `v1.0.0`.

**Architecture:** One facade class `DotNetVBA` reached through a public function `Net()` in a standard module; one `PublicNotCreatable` class per .NET static class, all late-bound over `Scripting.FileSystemObject`, `ADODB.Stream`, `WScript.Shell`, `Application.FileDialog` and a few `kernel32` `PtrSafe` declares. A PowerShell runner drives Excel and Word over COM, imports the sources plus a plain `.bas` test module, runs it, and reads a log file. A build script turns `src/` into `dist/DotNetVBA.xlam`.

**Tech Stack:** VBA 7.1 (64-bit Office 2016+), PowerShell 7 (COM automation), git. No external VBA frameworks or references.

**Spec:** `docs/superpowers/specs/2026-10-27-dotnetvba-design.md`

## Global Constraints

- Every module starts with `Option Explicit`. Follow the `vba-coding-conventions` skill (no identifiers named `Key`, `Name`, `Type`, `Count` as parameters, etc.).
- `Declare PtrSafe` only; no 32-bit branches.
- Late binding only: `CreateObject(...)`, `As Object`; Office constants inlined with their numeric values and a comment.
- Source bytes in `src/` and `tests/` are pure ASCII, no BOM, CRLF line endings. Non-ASCII test data is built with `ChrW`.
- Every class module header: `Attribute VB_Creatable = False`, `Attribute VB_Exposed = True` (PublicNotCreatable), `Attribute VB_PredeclaredId = False`.
- No Excel-only API anywhere in `src/` (`Application.GetOpenFilename`, `Range`, `Cells`, `ThisWorkbook`). `Application.FileDialog`, `Application.Name`, `Application.VBE` are allowed.
- Errors: `Err.Raise vbObjectError + n, "DotNetVBA.<Class>.<Member>", "<ExceptionName>: <message>"` with `n` = 1001 ArgumentException, 1002 ArgumentOutOfRangeException, 1003 FileNotFoundException, 1004 DirectoryNotFoundException, 1005 IOException, 1006 InvalidOperationException, 1007 NotSupportedException.
- Empty string-array results are `Split(vbNullString)` (`LBound = 0`, `UBound = -1`), never uninitialized.
- Text default encoding `"utf-8"`, written **without** BOM. `WriteAllLines` terminates every line with `vbCrLf`.
- `List` is 0-based; `Item` is the default member; `NewEnum` has `VB_UserMemId = -4`.
- Git identity: the machine's global config (`straygizmo`). Commit after every task. Branch `main`.
- Version string `"1.0.0"` lives in exactly one place: `Private Const VERSION_STRING As String = "1.0.0"` in `src/DotNetVBA.cls`.

## Review Focus

1. **Trailing backslash on directory paths** (`C:\tmp\dir\`): `Directory.Exists`, `GetFiles`, `Path.GetFileName` must behave like .NET (`GetFileName` returns `""`, `Exists` is True). Test pinned in Task 2 and Task 4.
2. **Non-ASCII (Japanese) file and folder names**: every `File`/`Directory` member must round-trip a name containing `ChrW(&H3042)`. Test pinned in Task 4 and Task 5.
3. **Empty-string arguments**: `File.Exists("")` and `Directory.Exists("")` return False and `Path.Combine("", "")` returns `""` (as .NET); `ReadAllText("")`, `Copy("", x)`, `IniFile.Read("", ...)` → `ArgumentException` (1001), never a raw FSO error. Tests pinned in Tasks 2, 4, 5, 8.
4. **Empty file and >32767 lines**: `ReadAllLines` on a 0-byte file returns an empty array; on a 40000-line file returns 40000 elements (no `Integer` overflow anywhere). Test pinned in Task 5.
5. **Mixed value/object lists**: `Contains`/`IndexOf`/`Remove` on a list holding both a string and an object never raise type-mismatch; `Sort` on such a list raises `InvalidOperationException`. Test pinned in Task 7.

---

## File structure

| Path | Responsibility |
|---|---|
| `src/DotNetVBAGlobals.bas` | `Public Function Net() As DotNetVBA` singleton accessor. Nothing else public. |
| `src/DotNetVBAInternal.bas` | `Option Private Module`. Shared helpers: `ThrowDotNet`, `NewFso`, `EmptyStringArray`, `WildcardToLike`, `DotNetErrorNumber` enum. Invisible to referencing projects. |
| `src/DotNetVBA.cls` | Facade: lazy properties for each class, `NewList`, `Version`. |
| `src/DotNetPath.cls` | Pure string path functions. |
| `src/DotNetDirectory.cls` | Folder operations over FSO. |
| `src/DotNetFile.cls` | File operations and text IO over FSO + ADODB.Stream. |
| `src/DotNetEnvironment.cls` | Special folders, env vars, `SpecialFolder` enum. |
| `src/DotNetList.cls` | 0-based Variant list with default member and `For Each`. |
| `src/DotNetIniFile.cls` | Win32 private-profile API. |
| `src/DotNetThread.cls` | `Sleep`. |
| `src/DotNetProcess.cls` | `Start`. |
| `src/DotNetDialog.cls` | Office `FileDialog` wrappers. |
| `tests/TestDotNetVBA.bas` | Assert helpers, `RunAllTests`, all `Test_*` procedures. |
| `tests/TestAddInConsumer.bas` | Macro run from a workbook that *references* the `.xlam`. |
| `tests/run-tests.ps1` | ASCII/CRLF lint + Excel/Word import-mode run + add-in run. |
| `tests/MANUAL.md` | Manual checklist for `Dialog` and `Process.Start`. |
| `build/build-xlam.ps1` | Builds `dist/DotNetVBA.xlam`. |
| `README.md`, `LICENSE`, `.gitignore`, `.gitattributes` | Repo metadata. |

Test module conventions used by every task (defined in Task 1):

```vb
' Each test is: Public Sub Test_<Class>_<Behavior>()
' Helpers available:
'   AssertEqual expected, actual, [message]        ' compares with = after CStr for strings, exact for numbers
'   AssertTrue  condition, [message]
'   AssertFalse condition, [message]
'   AssertArrayEqual expectedArray, actualArray    ' same bounds and elements
'   AssertRaised expectedErrorNumber               ' reads Err.Number, clears Err; call right after On Error Resume Next block
'   NewTempDir() As String                         ' fresh folder under Path.GetTempPath, deleted after the test
'   WriteBytes path, bytes()                       ' raw helper via ADODB.Stream binary
'   ReadBytes(path) As Byte()                      ' raw helper
```

---

### Task 1: Repository scaffold, test harness, facade skeleton

**Files:**
- Delete: the five pre-existing `*.cls` files in the repository root
- Create: `.gitignore`, `.gitattributes`, `LICENSE`
- Create: `src/DotNetVBAGlobals.bas`, `src/DotNetVBAInternal.bas`, `src/DotNetVBA.cls`
- Create: `tests/TestDotNetVBA.bas`, `tests/run-tests.ps1`

**Interfaces:**
- Produces: `Net() As DotNetVBA`; `DotNetVBA.Version As String`; `DotNetVBA.NewList()` is **not** yet present (Task 7).
- Produces (internal): `ThrowDotNet(errNumber As DotNetErrorNumber, source As String, message As String)`; `NewFso() As Object`; `EmptyStringArray() As String()`; `Public Enum DotNetErrorNumber: DotNetArgument = 1001, DotNetArgumentOutOfRange = 1002, DotNetFileNotFound = 1003, DotNetDirectoryNotFound = 1004, DotNetIO = 1005, DotNetInvalidOperation = 1006, DotNetNotSupported = 1007`. `ThrowDotNet` raises `vbObjectError + errNumber`.
- Produces (tests): the helper API listed in "File structure"; `RunAllTests(outputPath As String) As Long` returning the failure count and writing one line per test: `PASS <name>` or `FAIL <name>: <message>`, followed by a final line `TOTAL <n> FAILED <m>`.
- Produces (runner): `tests/run-tests.ps1 [-Office Excel|Word|All] [-SkipAddIn]`, exit code = failures.

- [ ] **Step 1: git init and delete the pre-existing files**

```powershell
git init -b main
Remove-Item .\*.cls      # the five pre-existing class files in the root; src\ is created fresh
```

- [ ] **Step 2: Write `.gitignore`, `.gitattributes`, `LICENSE`**

`.gitignore`: `dist/`, `~$*`, `*.xlsm`, `*.docm`, `*.log`, `tests/out/`.
`.gitattributes`: `*.cls text eol=crlf`, `*.bas text eol=crlf`, `*.ps1 text eol=crlf`, `*.md text eol=lf`.
`LICENSE`: MIT, `Copyright (c) 2026 straygizmo`.

- [ ] **Step 3: Write `src/DotNetVBAInternal.bas`**

Header `Attribute VB_Name = "DotNetVBAInternal"`, then `Option Explicit`, `Option Private Module`. Define the `DotNetErrorNumber` enum and:

```vb
Public Sub ThrowDotNet(ByVal errNumber As DotNetErrorNumber, ByVal source As String, ByVal message As String)
Public Function NewFso() As Object                 ' CreateObject("Scripting.FileSystemObject")
Public Function EmptyStringArray() As String()     ' Split(vbNullString)
Public Function ExceptionName(ByVal errNumber As DotNetErrorNumber) As String  ' "FileNotFoundException" etc.
```

`ThrowDotNet` builds the description as `ExceptionName(errNumber) & ": " & message` and the source as `"DotNetVBA." & source`.

- [ ] **Step 4: Write `src/DotNetVBA.cls` with only `Version`**

Class header per Global Constraints, `Private Const VERSION_STRING As String = "1.0.0"`, `Public Property Get Version() As String`.

- [ ] **Step 5: Write `src/DotNetVBAGlobals.bas`**

```vb
Private mInstance As DotNetVBA
Public Function Net() As DotNetVBA   ' creates mInstance on first call
```

- [ ] **Step 6: Write `tests/TestDotNetVBA.bas` with the harness and one smoke test**

Harness design:
- `RunAllTests(outputPath As String) As Long` opens the log with FSO (`CreateTextFile`, overwrite), enumerates procedures of its own module through `Application.VBE.ActiveVBProject.VBComponents("TestDotNetVBA").CodeModule` (scan `Lines(i, 1)` for lines starting with `Public Sub Test_`), and calls each with `Application.Run procName` inside `On Error Resume Next`; non-zero `Err.Number` after the call means FAIL with `Err.Description`. After each test it deletes the temp dir created by `NewTempDir` if any.
- Assert failures raise `vbObjectError + 9999` with the message; the runner logs them as FAIL.
- `AssertRaised(expectedNumber As Long)`: captures `Err.Number` and `Err.Description` first thing, calls `Err.Clear`, then fails if the number differs (message includes the captured description).

Smoke test:

```vb
Public Sub Test_Facade_Version()
    AssertEqual "1.0.0", Net.Version
    AssertTrue Net Is Net, "Net() must return the same instance"
End Sub
```

- [ ] **Step 7: Write `tests/run-tests.ps1`**

```powershell
param([ValidateSet('Excel','Word','All')][string]$Office = 'All', [switch]$SkipAddIn)
```

Stages:
1. **Lint**: for every file in `src\*.cls`, `src\*.bas`, `tests\*.bas`: read bytes; fail if any byte ≥ 0x80, if the file starts with EF BB BF, or if any LF is not preceded by CR. Print `LINT OK (<n> files)`.
2. **Import-mode run** (`Invoke-ImportModeRun -App 'Excel'|'Word'`): `New-Object -ComObject Excel.Application` / `Word.Application`; `Visible=$false`, `DisplayAlerts=$false` (Excel) / `DisplayAlerts=0` (Word); `Workbooks.Add()` / `Documents.Add()`; `$proj = $doc.VBProject` wrapped in try/catch that prints `Enable "Trust access to the VBA project object model" in <App> Trust Center` and exits 2; import every `src\*.bas`, `src\*.cls`, then `tests\TestDotNetVBA.bas` via `$proj.VBComponents.Import($path)`; `$log = Join-Path $env:TEMP "dotnetvba-$App.log"`; `$failures = $app.Run('RunAllTests', $log)` (if Word rejects the bare name, use `'TestDotNetVBA.RunAllTests'`); print the log; close without saving (`$doc.Close($false)` / `$doc.Close(0)`); `$app.Quit()`; release with `[System.Runtime.InteropServices.Marshal]::ReleaseComObject`.
3. **Add-in run**: placeholder that prints `ADDIN: skipped (not built yet)` until Task 11 replaces it.
4. Exit with the sum of failures.

- [ ] **Step 8: Run the runner and verify the smoke test passes in both hosts**

Run: `pwsh -File tests\run-tests.ps1`
Expected: `LINT OK (3 files)`, then for Excel and Word a log containing `PASS Test_Facade_Version` and `TOTAL 1 FAILED 0`; exit code 0.

- [ ] **Step 9: Commit**

```bash
git add -A
git commit -m "chore: scaffold DotNetVBA with test harness and facade skeleton"
```

---

### Task 2: DotNetPath

**Files:**
- Create: `src/DotNetPath.cls`
- Modify: `src/DotNetVBA.cls` (add `Path` property)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.Path` of type `DotNetPath` with members exactly as spec §6.4: `Combine(ParamArray parts() As Variant) As String`, `GetFileName`, `GetFileNameWithoutExtension`, `GetExtension`, `GetDirectoryName`, `ChangeExtension(ByVal path As String, ByVal extension As String)`, `HasExtension`, `IsPathRooted`, `GetPathRoot`, `GetFullPath`, `GetTempPath`, `GetTempFileName`, `GetRandomFileName`, `DirectorySeparatorChar` (Property Get).
- Consumes: `ThrowDotNet`, `NewFso`.

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_Path_Combine()
    AssertEqual "C:\a\b\c.txt", Net.Path.Combine("C:\a", "b", "c.txt")
    AssertEqual "C:\a\b", Net.Path.Combine("C:\a\", "b")
    AssertEqual "D:\x", Net.Path.Combine("C:\a", "D:\x")       ' rooted later part wins
    AssertEqual "C:\a\b", Net.Path.Combine("C:\a", "", "b")    ' empty parts skipped
    AssertEqual "", Net.Path.Combine("", "")                      ' all empty -> "" (as .NET)
    AssertEqual "b", Net.Path.Combine("", "b")
End Sub

Public Sub Test_Path_FileNameParts()
    AssertEqual "c.txt", Net.Path.GetFileName("C:\a\b\c.txt")
    AssertEqual "", Net.Path.GetFileName("C:\a\b\")
    AssertEqual "c", Net.Path.GetFileNameWithoutExtension("C:\a\b\c.txt")
    AssertEqual ".txt", Net.Path.GetExtension("C:\a\b\c.txt")
    AssertEqual "", Net.Path.GetExtension("C:\a\b\c")
    AssertEqual ".gz", Net.Path.GetExtension("x.tar.gz")
    AssertEqual "C:\a\b", Net.Path.GetDirectoryName("C:\a\b\c.txt")
    AssertEqual "C:\a\b", Net.Path.GetDirectoryName("C:\a\b\")
    AssertTrue Net.Path.HasExtension("a.b")
    AssertFalse Net.Path.HasExtension("a")
End Sub

Public Sub Test_Path_ChangeExtension()
    AssertEqual "C:\a\c.md", Net.Path.ChangeExtension("C:\a\c.txt", ".md")
    AssertEqual "C:\a\c.md", Net.Path.ChangeExtension("C:\a\c.txt", "md")
    AssertEqual "C:\a\c", Net.Path.ChangeExtension("C:\a\c.txt", "")
    AssertEqual "C:\a\c.md", Net.Path.ChangeExtension("C:\a\c", ".md")
End Sub

Public Sub Test_Path_Roots()
    AssertTrue Net.Path.IsPathRooted("C:\a")
    AssertTrue Net.Path.IsPathRooted("\\server\share\a")
    AssertFalse Net.Path.IsPathRooted("a\b")
    AssertEqual "C:\", Net.Path.GetPathRoot("C:\a\b")
    AssertEqual "\\server\share", Net.Path.GetPathRoot("\\server\share\a")
    AssertEqual "", Net.Path.GetPathRoot("a\b")
    AssertEqual "\", Net.Path.DirectorySeparatorChar
End Sub

Public Sub Test_Path_FullAndTemp()
    Dim tmp As String
    tmp = Net.Path.GetTempPath()
    AssertEqual "\", Right$(tmp, 1)
    AssertEqual "C:\a\c", Net.Path.GetFullPath("C:\a\b\..\c")
    Dim f As String
    f = Net.Path.GetTempFileName()
    AssertTrue NewFso().FileExists(f), "GetTempFileName must create the file"
    NewFso().DeleteFile f
    AssertTrue Len(Net.Path.GetRandomFileName()) >= 8
    AssertTrue Net.Path.GetRandomFileName() <> Net.Path.GetRandomFileName()
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: log shows `FAIL Test_Path_*` with a "Method or data member not found"/compile error on `Net.Path`. (If the module fails to compile at all, every test fails — that is the expected signal here.)

- [ ] **Step 3: Implement `src/DotNetPath.cls` and the `Path` property on the facade**

Pure `InStrRev`/`Mid$` string work; `GetFullPath` = `NewFso().GetAbsolutePathName(path)`; `GetTempPath` = `NewFso().GetSpecialFolder(2).Path & "\"`; `GetTempFileName` creates a zero-byte file named `NewFso().GetTempName()` in the temp folder; `GetRandomFileName` returns `NewFso().GetTempName()` without creating anything. `Combine`: skip empty parts; all parts empty returns `""` (as .NET); a part that `IsPathRooted` resets the accumulator. `IsPathRooted`: second char is `:` or starts with `\`. Facade property is a lazy singleton: `If mPath Is Nothing Then Set mPath = New DotNetPath`.

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_Path_*` PASS in Excel and Word; `FAILED 0`.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetPath.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetPath (System.IO.Path)"
```

---

### Task 3: Test-side file helpers and temp-dir fixture

**Files:**
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `NewTempDir() As String` (creates `Net.Path.GetTempPath() & "dotnetvba-" & NewFso().GetTempName()`, records it in a module variable; `RunAllTests` deletes it after the test with `DeleteFolder path, True` under `On Error Resume Next`); `WriteBytes(ByVal path As String, bytes() As Byte)` and `ReadBytes(ByVal path As String) As Byte()` through `ADODB.Stream` `Type = 1`; `WriteTextRaw(ByVal path As String, ByVal text As String, ByVal charset As String)` through `ADODB.Stream` text mode (BOM behavior is whatever ADODB does — used only to create fixtures).

- [ ] **Step 1: Write a test for the fixture itself**

```vb
Public Sub Test_Harness_TempDirAndBytes()
    Dim dir As String
    dir = NewTempDir()
    AssertTrue NewFso().FolderExists(dir)
    Dim b(0 To 2) As Byte
    b(0) = 65: b(1) = 66: b(2) = 67
    WriteBytes dir & "\x.bin", b
    Dim r() As Byte
    r = ReadBytes(dir & "\x.bin")
    AssertEqual 2, UBound(r)
    AssertEqual 67, r(2)
End Sub
```

- [ ] **Step 2: Run, verify it fails (helpers missing), implement the helpers, run again**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected after implementation: `PASS Test_Harness_TempDirAndBytes`, and the temp folder no longer exists after the run (check `Get-ChildItem $env:TEMP dotnetvba-*` returns nothing).

- [ ] **Step 3: Commit**

```bash
git add tests/TestDotNetVBA.bas
git commit -m "test: add temp-dir fixture and byte helpers"
```

---

### Task 4: DotNetDirectory

**Files:**
- Create: `src/DotNetDirectory.cls`
- Modify: `src/DotNetVBA.cls` (add `Directory` property)
- Modify: `src/DotNetVBAInternal.bas` (add `WildcardToLike`)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.Directory` of type `DotNetDirectory` with members as spec §6.3: `Exists`, `CreateDirectory`, `Delete(ByVal path As String, Optional ByVal recursive As Boolean = False)`, `Move`, `GetFiles(ByVal path As String, Optional ByVal searchPattern As String = "*", Optional ByVal recursive As Boolean = False) As String()`, `GetDirectories(...same...)`, `GetParent`, `GetCurrentDirectory`, `SetCurrentDirectory`, `GetCreationTime`, `GetLastWriteTime`.
- Produces (internal): `WildcardToLike(ByVal pattern As String) As String` (escapes `[` → `[[]`, `#` → `[#]`, leaves `*` and `?`), `MatchesPattern(ByVal nameText As String, ByVal pattern As String) As Boolean` (`LCase$(nameText) Like LCase$(WildcardToLike(pattern))`).

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_Directory_CreateExistsDelete()
    Dim dir As String: dir = NewTempDir()
    AssertFalse Net.Directory.Exists("")
    AssertTrue Net.Directory.Exists(dir)
    AssertTrue Net.Directory.Exists(dir & "\")                       ' trailing backslash
    Net.Directory.CreateDirectory dir & "\a\b\c"                     ' recursive create
    AssertTrue Net.Directory.Exists(dir & "\a\b\c")
    Net.Directory.CreateDirectory dir & "\a\b\c"                     ' no error if exists
    Net.Directory.Delete dir & "\a\b\c"                              ' empty, non-recursive OK
    AssertFalse Net.Directory.Exists(dir & "\a\b\c")
    On Error Resume Next
    Net.Directory.Delete dir & "\a"                                  ' non-empty
    AssertRaised vbObjectError + 1005
    Net.Directory.Delete dir & "\missing"
    AssertRaised vbObjectError + 1004
    On Error GoTo 0
    Net.Directory.Delete dir & "\a", True
    AssertFalse Net.Directory.Exists(dir & "\a")
End Sub

Public Sub Test_Directory_NonAsciiNames()
    Dim dir As String: dir = NewTempDir()
    Dim jp As String: jp = dir & "\" & ChrW(&H3042) & ChrW(&H3044)
    Net.Directory.CreateDirectory jp
    AssertTrue Net.Directory.Exists(jp)
    Dim dirs() As String: dirs = Net.Directory.GetDirectories(dir)
    AssertEqual 0, UBound(dirs)
    AssertEqual jp, dirs(0)
End Sub

Public Sub Test_Directory_GetFilesAndDirectories()
    Dim dir As String: dir = NewTempDir()
    Dim fso As Object: Set fso = NewFso()
    fso.CreateTextFile(dir & "\a.txt").Close
    fso.CreateTextFile(dir & "\b.log").Close
    fso.CreateFolder dir & "\sub"
    fso.CreateTextFile(dir & "\sub\c.txt").Close
    Dim r() As String
    r = Net.Directory.GetFiles(dir)
    AssertEqual 1, UBound(r)                                         ' full paths, non-recursive
    AssertEqual dir & "\a.txt", r(0)
    r = Net.Directory.GetFiles(dir, "*.txt")
    AssertEqual 0, UBound(r)
    r = Net.Directory.GetFiles(dir, "*.TXT", True)                   ' case-insensitive, recursive
    AssertEqual 1, UBound(r)
    AssertEqual dir & "\sub\c.txt", r(1)
    r = Net.Directory.GetFiles(dir, "?.log")
    AssertEqual 0, UBound(r)
    r = Net.Directory.GetFiles(dir, "*.none")
    AssertEqual -1, UBound(r)                                        ' empty array, not error
    r = Net.Directory.GetDirectories(dir)
    AssertEqual dir & "\sub", r(0)
    r = Net.Directory.GetFiles(dir & "\")                            ' trailing backslash: no double separator
    AssertEqual dir & "\a.txt", r(0)
    On Error Resume Next
    r = Net.Directory.GetFiles(dir & "\missing")
    AssertRaised vbObjectError + 1004
    On Error GoTo 0
End Sub

Public Sub Test_Directory_MoveParentCurrent()
    Dim dir As String: dir = NewTempDir()
    Net.Directory.CreateDirectory dir & "\src"
    Net.Directory.Move dir & "\src", dir & "\dst"
    AssertTrue Net.Directory.Exists(dir & "\dst")
    AssertFalse Net.Directory.Exists(dir & "\src")
    AssertEqual dir, Net.Directory.GetParent(dir & "\dst")
    AssertEqual "", Net.Directory.GetParent("C:\")
    Dim saved As String: saved = Net.Directory.GetCurrentDirectory()
    Net.Directory.SetCurrentDirectory dir
    AssertEqual dir, Net.Directory.GetCurrentDirectory()
    Net.Directory.SetCurrentDirectory saved
    AssertTrue Abs(Now - Net.Directory.GetLastWriteTime(dir & "\dst")) < 1 / 24
    AssertTrue Abs(Now - Net.Directory.GetCreationTime(dir & "\dst")) < 1 / 24
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: `FAIL Test_Directory_*`.

- [ ] **Step 3: Implement `src/DotNetDirectory.cls`, `WildcardToLike`, and the facade property**

Notes that the signature does not settle: strip one trailing `\` from `path` at the top of every member (except for roots like `C:\`); `CreateDirectory` recurses on `NewFso().GetParentFolderName`; `Delete` non-recursive checks `Files.Count = 0 And SubFolders.Count = 0` else `DotNetIO`; `GetFiles` builds results in a `Collection`, then copies to a `String()`; results are `folder.Path & "\" & item.Name` so the trailing-backslash case cannot double up; `GetCurrentDirectory` = `CurDir$()`, `SetCurrentDirectory` = `ChDrive` + `ChDir`.

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_Directory_*` PASS in Excel and Word.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetDirectory.cls src/DotNetVBA.cls src/DotNetVBAInternal.bas tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetDirectory (System.IO.Directory)"
```

---

### Task 5: DotNetFile

**Files:**
- Create: `src/DotNetFile.cls`
- Modify: `src/DotNetVBA.cls` (add `File` property)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.File` of type `DotNetFile` with members as spec §6.2: `Exists`, `Copy(ByVal source As String, ByVal dest As String, Optional ByVal overwrite As Boolean = False)`, `Move(...same...)`, `Delete`, `ReadAllText(ByVal path As String, Optional ByVal encoding As String = "utf-8") As String`, `ReadAllLines(...) As String()`, `WriteAllText(ByVal path As String, ByVal contents As String, Optional ByVal encoding As String = "utf-8")`, `WriteAllLines(ByVal path As String, ByVal lines As Variant, Optional ByVal encoding As String = "utf-8")`, `AppendAllText`, `AppendAllLines`, `GetCreationTime`, `GetLastWriteTime`, `GetLastAccessTime`, `GetLength(ByVal path As String) As Double`, `GetAttributes(ByVal path As String) As Long`.
- `lines` accepts a `String()`/`Variant()` array or a `DotNetList` (Task 7 adds the list; until then, only arrays are tested — the `DotNetList` branch is added in Task 7).

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_File_ExistsCopyMoveDelete()
    Dim dir As String: dir = NewTempDir()
    Dim p As String: p = dir & "\a.txt"
    AssertFalse Net.File.Exists("")
    AssertFalse Net.File.Exists(p)
    Net.File.WriteAllText p, "hello"
    AssertTrue Net.File.Exists(p)
    AssertFalse Net.File.Exists(dir)                                 ' folder is not a file
    Net.File.Copy p, dir & "\b.txt"
    AssertEqual "hello", Net.File.ReadAllText(dir & "\b.txt")
    On Error Resume Next
    Net.File.Copy p, dir & "\b.txt"
    AssertRaised vbObjectError + 1005                                ' exists, no overwrite
    Net.File.Copy "", dir & "\c.txt"
    AssertRaised vbObjectError + 1001
    Net.File.Copy dir & "\missing.txt", dir & "\c.txt"
    AssertRaised vbObjectError + 1003
    On Error GoTo 0
    Net.File.WriteAllText p, "new"
    Net.File.Copy p, dir & "\b.txt", True
    AssertEqual "new", Net.File.ReadAllText(dir & "\b.txt")
    Net.File.Move dir & "\b.txt", dir & "\c.txt"
    AssertFalse Net.File.Exists(dir & "\b.txt")
    AssertTrue Net.File.Exists(dir & "\c.txt")
    Net.File.Move p, dir & "\c.txt", True
    AssertEqual "new", Net.File.ReadAllText(dir & "\c.txt")
    Net.File.Delete dir & "\c.txt"
    Net.File.Delete dir & "\c.txt"                                   ' no error when missing
    AssertFalse Net.File.Exists(dir & "\c.txt")
End Sub

Public Sub Test_File_Utf8NoBom()
    Dim dir As String: dir = NewTempDir()
    Dim p As String: p = dir & "\u.txt"
    Dim s As String: s = "A" & ChrW(&H3042) & "B"
    Net.File.WriteAllText p, s
    Dim b() As Byte: b = ReadBytes(p)
    AssertEqual 4, UBound(b)                                         ' 'A' + 3 bytes + 'B' = 5 bytes, no BOM
    AssertEqual 65, b(0)
    AssertEqual s, Net.File.ReadAllText(p)
    Dim bom(0 To 3) As Byte: bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF: bom(3) = 65
    WriteBytes p, bom
    AssertEqual "A", Net.File.ReadAllText(p)                         ' BOM tolerated on read
End Sub

Public Sub Test_File_OtherEncodings()
    Dim dir As String: dir = NewTempDir()
    Dim s As String: s = "A" & ChrW(&H3042)
    Net.File.WriteAllText dir & "\sj.txt", s, "shift_jis"
    AssertEqual 2, UBound(ReadBytes(dir & "\sj.txt"))                ' 1 + 2 bytes
    AssertEqual s, Net.File.ReadAllText(dir & "\sj.txt", "shift_jis")
    Net.File.WriteAllText dir & "\u16.txt", s, "utf-16"
    AssertEqual s, Net.File.ReadAllText(dir & "\u16.txt", "utf-16")
End Sub

Public Sub Test_File_Lines()
    Dim dir As String: dir = NewTempDir()
    Dim p As String: p = dir & "\l.txt"
    Dim want(0 To 2) As String: want(0) = "a": want(1) = "": want(2) = "c"
    Net.File.WriteAllLines p, want
    AssertEqual "a" & vbCrLf & vbCrLf & "c" & vbCrLf, Net.File.ReadAllText(p)
    AssertArrayEqual want, Net.File.ReadAllLines(p)
    WriteTextRaw p, "a" & vbLf & vbLf & "c", "utf-8"                 ' LF, no trailing newline
    AssertArrayEqual want, Net.File.ReadAllLines(p)
    WriteTextRaw p, "a" & vbCr & vbCr & "c" & vbCr, "utf-8"          ' CR
    AssertArrayEqual want, Net.File.ReadAllLines(p)
    WriteTextRaw p, "", "utf-8"
    AssertEqual -1, UBound(Net.File.ReadAllLines(p))                 ' empty file
    Net.File.AppendAllText p, "x"
    Net.File.AppendAllText p, "y"
    AssertEqual "xy", Net.File.ReadAllText(p)
    Dim more(0 To 0) As String: more(0) = "z"
    Net.File.AppendAllLines p, more
    AssertEqual "xyz" & vbCrLf, Net.File.ReadAllText(p)
    Net.File.AppendAllText dir & "\new.txt", "q"                     ' creates when missing
    AssertEqual "q", Net.File.ReadAllText(dir & "\new.txt")
End Sub

Public Sub Test_File_ManyLines()
    Dim dir As String: dir = NewTempDir()
    Dim i As Long, arr(0 To 39999) As String
    For i = 0 To 39999: arr(i) = CStr(i): Next
    Net.File.WriteAllLines dir & "\big.txt", arr
    Dim r() As String: r = Net.File.ReadAllLines(dir & "\big.txt")
    AssertEqual 39999, UBound(r)
    AssertEqual "39999", r(39999)
End Sub

Public Sub Test_File_ErrorsAndInfo()
    Dim dir As String: dir = NewTempDir()
    On Error Resume Next
    Net.File.ReadAllText dir & "\missing.txt"
    AssertRaised vbObjectError + 1003
    Net.File.ReadAllText ""
    AssertRaised vbObjectError + 1001
    Net.File.WriteAllText dir & "\nodir\x.txt", "x"
    AssertRaised vbObjectError + 1004
    On Error GoTo 0
    Dim jp As String: jp = dir & "\" & ChrW(&H3042) & ".txt"
    Net.File.WriteAllText jp, "12345"
    AssertEqual 5, Net.File.GetLength(jp)
    AssertTrue Abs(Now - Net.File.GetLastWriteTime(jp)) < 1 / 24
    AssertTrue Abs(Now - Net.File.GetCreationTime(jp)) < 1 / 24
    AssertTrue Abs(Now - Net.File.GetLastAccessTime(jp)) < 1
    AssertEqual 0, Net.File.GetAttributes(jp) And 16                 ' not a directory bit
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: `FAIL Test_File_*`.

- [ ] **Step 3: Implement `src/DotNetFile.cls` and the facade property**

Decisions the tests do not settle:
- Argument check: empty `path`/`source`/`dest` → `DotNetArgument`; missing source file → `DotNetFileNotFound`; missing parent folder on write → `DotNetDirectoryNotFound`.
- Read: `ADODB.Stream` `Type = 2`, `Charset = encoding`, `Open`, `LoadFromFile`, `ReadText(-1)`. ADODB skips a UTF-8 BOM itself.
- Write without BOM (the one algorithm the signature does not determine):
  1. text stream: `Type = 2`, `Charset = encoding`, `Open`, `WriteText contents`;
  2. `Position = 0`, then `Type = 1` (Type may only change at position 0);
  3. if `LCase$(encoding) = "utf-8"` and `Size >= 3` and the first three bytes read are `EF BB BF`, set `Position = 3`; else `Position = 0`;
  4. new binary stream `Type = 1`, `Open`; `textStream.CopyTo binaryStream`; `binaryStream.SaveToFile path, 2` (`adSaveCreateOverWrite = 2`).
  Other encodings keep whatever BOM ADODB emits (utf-16 gets FF FE, matching .NET's `Encoding.Unicode`).
- `ReadAllLines`: `ReadAllText`, replace `vbCrLf` → `vbLf`, `vbCr` → `vbLf`, drop exactly one trailing `vbLf`, return `EmptyStringArray()` if the result is empty, else `Split(text, vbLf)`.
- `WriteAllLines`: join elements with `vbCrLf` and add a trailing `vbCrLf`; accept any array (`IsArray`) now; Task 7 adds `TypeName(lines) = "DotNetList"` → `lines.ToArray()`.
- `AppendAllText`: if the file exists, `WriteAllText path, ReadAllText(path, encoding) & contents, encoding`; else `WriteAllText`.
- `GetLength` = `fso.GetFile(path).Size` as `Double`; dates from the FSO `File` object; `GetAttributes` = `.Attributes`.

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_File_*` PASS in Excel and Word.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetFile.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetFile (System.IO.File) with BOM-less UTF-8"
```

---

### Task 6: DotNetEnvironment

**Files:**
- Create: `src/DotNetEnvironment.cls`
- Modify: `src/DotNetVBA.cls` (add `Environment` property)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.Environment` of type `DotNetEnvironment`; `Public Enum SpecialFolder` with the twelve `SpecialFolder_*` members from spec §6.5 declared in this class module; `GetFolderPath(ByVal folder As SpecialFolder) As String`, `NewLine`, `UserName`, `MachineName`, `CurrentDirectory` (Get/Let), `GetEnvironmentVariable(ByVal variableName As String) As String`, `ExpandEnvironmentVariables(ByVal text As String) As String`, `Is64BitProcess`.

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_Environment_Folders()
    Dim d As String: d = Net.Environment.GetFolderPath(SpecialFolder_Desktop)
    AssertTrue Net.Directory.Exists(d), "Desktop: " & d
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_MyDocuments))
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_ApplicationData))
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_LocalApplicationData))
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_UserProfile))
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_ProgramFiles))
    AssertTrue Net.Directory.Exists(Net.Environment.GetFolderPath(SpecialFolder_Templates))
    On Error Resume Next
    Net.Environment.GetFolderPath 999
    AssertRaised vbObjectError + 1001
    On Error GoTo 0
End Sub

Public Sub Test_Environment_Misc()
    AssertEqual vbCrLf, Net.Environment.NewLine
    AssertTrue Len(Net.Environment.UserName) > 0
    AssertTrue Len(Net.Environment.MachineName) > 0
    AssertEqual Environ$("USERNAME"), Net.Environment.UserName
    AssertEqual "", Net.Environment.GetEnvironmentVariable("DOTNETVBA_NOT_SET_12345")
    AssertEqual Environ$("TEMP"), Net.Environment.GetEnvironmentVariable("TEMP")
    AssertEqual Environ$("TEMP") & "\x", Net.Environment.ExpandEnvironmentVariables("%TEMP%\x")
    AssertTrue Net.Environment.Is64BitProcess
    Dim dir As String: dir = NewTempDir()
    Dim saved As String: saved = Net.Environment.CurrentDirectory
    Net.Environment.CurrentDirectory = dir
    AssertEqual dir, Net.Environment.CurrentDirectory
    Net.Environment.CurrentDirectory = saved
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: `FAIL Test_Environment_*`.

- [ ] **Step 3: Implement `src/DotNetEnvironment.cls` and the facade property**

`GetFolderPath` maps: Desktop/MyDocuments/Startup/StartMenu/Templates/Favorites/Recent/SendTo → `WScript.Shell.SpecialFolders("<name>")`; ApplicationData → `%APPDATA%`, LocalApplicationData → `%LOCALAPPDATA%`, UserProfile → `%USERPROFILE%`, ProgramFiles → `%ProgramFiles%` via `ExpandEnvironmentStrings`. Unknown enum value → `DotNetArgument`. `UserName`/`MachineName` from `Environ$`. `CurrentDirectory` delegates to `Net.Directory` semantics (`CurDir$`, `ChDrive`/`ChDir`).

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_Environment_*` PASS.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetEnvironment.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetEnvironment (System.Environment)"
```

---

### Task 7: DotNetList

**Files:**
- Create: `src/DotNetList.cls`
- Modify: `src/DotNetVBA.cls` (add `NewList() As DotNetList`)
- Modify: `src/DotNetFile.cls` (`WriteAllLines`/`AppendAllLines` accept `DotNetList`)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `DotNetList` with members as spec §6.6. Exact signatures: `Count As Long` (Get); `Item(ByVal index As Long) As Variant` Get/Let/Set with `Attribute Item.VB_UserMemId = 0`; `Add(ByVal item As Variant)`; `AddRange(ByVal items As Variant)`; `Insert(ByVal index As Long, ByVal item As Variant)`; `Remove(ByVal item As Variant) As Boolean`; `RemoveAt(ByVal index As Long)`; `RemoveRange(ByVal index As Long, ByVal itemCount As Long)`; `Clear()`; `Contains(ByVal item As Variant) As Boolean`; `IndexOf(ByVal item As Variant) As Long`; `LastIndexOf(ByVal item As Variant) As Long`; `GetRange(ByVal index As Long, ByVal itemCount As Long) As DotNetList`; `Sort()`; `Reverse()`; `ToArray() As Variant()`; `Join(Optional ByVal separator As String = ",") As String`; `NewEnum() As IUnknown` with `Attribute NewEnum.VB_UserMemId = -4` and `Attribute NewEnum.VB_MemberFlags = "40"`.
- Attribute lines go immediately after the `Public Property Get ...` line, inside the procedure.

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_List_AddItemCount()
    Dim l As DotNetList: Set l = Net.NewList()
    AssertEqual 0, l.Count
    l.Add "a": l.Add 2: l.Add 3.5
    AssertEqual 3, l.Count
    AssertEqual "a", l(0)                                            ' default member
    AssertEqual 2, l.Item(1)
    l(1) = 20
    AssertEqual 20, l(1)
    Dim o As Object: Set o = Net.NewList()
    Set l(2) = o
    AssertTrue l(2) Is o
    On Error Resume Next
    Dim v As Variant: v = l(3)
    AssertRaised vbObjectError + 1002
    v = l(-1)
    AssertRaised vbObjectError + 1002
    On Error GoTo 0
End Sub

Public Sub Test_List_InsertRemove()
    Dim l As DotNetList: Set l = Net.NewList()
    l.Add "b": l.Insert 0, "a": l.Insert 2, "c"                      ' insert at Count appends
    AssertEqual "a,b,c", l.Join()
    l.RemoveAt 1
    AssertEqual "a,c", l.Join()
    AssertTrue l.Remove("c")
    AssertFalse l.Remove("zzz")
    AssertEqual "a", l.Join()
    Dim arr(0 To 2) As Variant: arr(0) = 1: arr(1) = 2: arr(2) = 3
    l.AddRange arr
    AssertEqual "a,1,2,3", l.Join()
    l.RemoveRange 1, 2
    AssertEqual "a,3", l.Join()
    Dim other As DotNetList: Set other = Net.NewList()
    other.Add "x": other.Add "y"
    l.AddRange other
    AssertEqual "a|3|x|y", l.Join("|")
    l.Clear
    AssertEqual 0, l.Count
    On Error Resume Next
    l.Insert 1, "q"
    AssertRaised vbObjectError + 1002
    l.RemoveAt 0
    AssertRaised vbObjectError + 1002
    On Error GoTo 0
End Sub

Public Sub Test_List_Search()
    Dim l As DotNetList: Set l = Net.NewList()
    Dim o As Object: Set o = Net.NewList()
    l.Add "a": l.Add o: l.Add "a": l.Add 5
    AssertTrue l.Contains("a")
    AssertTrue l.Contains(o)
    AssertFalse l.Contains("5")                                      ' string "5" <> number 5
    AssertTrue l.Contains(5)
    AssertEqual 0, l.IndexOf("a")
    AssertEqual 2, l.LastIndexOf("a")
    AssertEqual 1, l.IndexOf(o)
    AssertEqual -1, l.IndexOf("none")
    AssertTrue l.Remove(o)
    AssertEqual 3, l.Count
End Sub

Public Sub Test_List_SortReverseRange()
    Dim l As DotNetList: Set l = Net.NewList()
    l.Add "b": l.Add "a": l.Add "B"
    l.Sort
    AssertEqual "B,a,b", l.Join()                                    ' binary compare
    l.Reverse
    AssertEqual "b,a,B", l.Join()
    Dim n As DotNetList: Set n = Net.NewList()
    n.Add 10: n.Add 9: n.Add 100
    n.Sort
    AssertEqual "9,10,100", n.Join()
    Dim r As DotNetList: Set r = n.GetRange(1, 2)
    AssertEqual "10,100", r.Join()
    AssertEqual 3, n.Count
    Dim arr() As Variant: arr = n.ToArray()
    AssertEqual 0, LBound(arr)
    AssertEqual 2, UBound(arr)
    AssertEqual 100, arr(2)
    Dim e As DotNetList: Set e = Net.NewList()
    arr = e.ToArray()
    AssertEqual -1, UBound(arr)
    AssertEqual "", e.Join()
    n.Add Net.NewList()
    On Error Resume Next
    n.Sort
    AssertRaised vbObjectError + 1006
    Set r = n.GetRange(2, 5)
    AssertRaised vbObjectError + 1002
    On Error GoTo 0
End Sub

Public Sub Test_List_ForEachAndGrowth()
    Dim l As DotNetList: Set l = Net.NewList()
    Dim i As Long
    For i = 1 To 1000: l.Add i: Next
    AssertEqual 1000, l.Count
    Dim total As Double, v As Variant
    For Each v In l
        total = total + v
    Next
    AssertEqual 500500, total
End Sub

Public Sub Test_List_WriteAllLinesAcceptsList()
    Dim dir As String: dir = NewTempDir()
    Dim l As DotNetList: Set l = Net.NewList()
    l.Add "x": l.Add "y"
    Net.File.WriteAllLines dir & "\l.txt", l
    AssertEqual "x" & vbCrLf & "y" & vbCrLf, Net.File.ReadAllText(dir & "\l.txt")
    Net.File.AppendAllLines dir & "\l.txt", l
    AssertEqual 3, UBound(Net.File.ReadAllLines(dir & "\l.txt"))
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: `FAIL Test_List_*`.

- [ ] **Step 3: Implement `src/DotNetList.cls`, `NewList` on the facade, and the `DotNetList` branch in `DotNetFile`**

Decisions the tests do not settle: backing store `Private mItems() As Variant`, `Private mCount As Long`, capacity starts at 4 and doubles (`ReDim Preserve`); element assignment goes through one private `SetAt(index, value)` that uses `Set` when `IsObject(value)`; equality through one private `ItemsEqual(a, b) As Boolean` (both objects → `a Is b`; one object → False; else `a = b` guarded against type mismatch with `On Error`; also require `VarType(a) = VarType(b)` when either is a string so `"5" <> 5`); `Sort` uses any O(n log n) algorithm with a private `CompareItems(a, b) As Long` (`StrComp(a, b, vbBinaryCompare)` when both are strings; numeric `<`/`>` otherwise) and raises `DotNetInvalidOperation` first if any item `IsObject`; `NewEnum` copies items into a fresh `Collection` and returns `col.[_NewEnum]`; `ToArray` on an empty list returns `Array()` (a `Variant()` with `UBound = -1`).

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_List_*` PASS in Excel and Word.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetList.cls src/DotNetVBA.cls src/DotNetFile.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetList (List<T>) with default member and For Each"
```

---

### Task 8: DotNetIniFile

**Files:**
- Create: `src/DotNetIniFile.cls`
- Modify: `src/DotNetVBA.cls` (add `IniFile` property)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.IniFile` with `Read(ByVal path As String, ByVal section As String, ByVal keyName As String, Optional ByVal defaultValue As String = "") As String`, `Write(ByVal path As String, ByVal section As String, ByVal keyName As String, ByVal value As String)`, `DeleteKey(ByVal path As String, ByVal section As String, ByVal keyName As String)`, `DeleteSection(ByVal path As String, ByVal section As String)`, `GetSections(ByVal path As String) As String()`, `GetKeys(ByVal path As String, ByVal section As String) As String()`.
- Declares (private to the class):

```vb
Private Declare PtrSafe Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpDefault As LongPtr, ByVal lpReturnedString As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function GetPrivateProfileSectionNamesW Lib "kernel32" (ByVal lpszReturnBuffer As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
```

- [ ] **Step 1: Write the failing tests**

```vb
Public Sub Test_IniFile_ReadWrite()
    Dim dir As String: dir = NewTempDir()
    Dim p As String: p = dir & "\c.ini"
    AssertEqual "dflt", Net.IniFile.Read(p, "S", "k", "dflt")       ' missing file -> default
    Net.IniFile.Write p, "S", "k", "v1"                              ' creates file
    AssertTrue Net.File.Exists(p)
    AssertEqual "v1", Net.IniFile.Read(p, "S", "k")
    AssertEqual "v1", Net.IniFile.Read(p, "s", "K")                  ' case-insensitive
    Net.IniFile.Write p, "S", "k", "v2"
    AssertEqual "v2", Net.IniFile.Read(p, "S", "k")
    Net.IniFile.Write p, "S", "k2", ""
    AssertEqual "", Net.IniFile.Read(p, "S", "k2", "dflt")           ' blank value is "", not default
    Net.IniFile.Write p, "T", "x", "1"
    AssertArrayEqual Split("S,T", ","), Net.IniFile.GetSections(p)
    AssertArrayEqual Split("k,k2", ","), Net.IniFile.GetKeys(p, "S")
    Net.IniFile.DeleteKey p, "S", "k2"
    AssertArrayEqual Split("k", ","), Net.IniFile.GetKeys(p, "S")
    Net.IniFile.DeleteSection p, "T"
    AssertArrayEqual Split("S", ","), Net.IniFile.GetSections(p)
    AssertEqual -1, UBound(Net.IniFile.GetKeys(p, "missing"))
    On Error Resume Next
    Net.IniFile.Write dir & "\nodir\c.ini", "S", "k", "v"
    AssertRaised vbObjectError + 1004
    Net.IniFile.Read "", "S", "k"
    AssertRaised vbObjectError + 1001
    On Error GoTo 0
End Sub

Public Sub Test_IniFile_RelativePathIsResolved()
    Dim dir As String: dir = NewTempDir()
    Dim saved As String: saved = Net.Directory.GetCurrentDirectory()
    Net.Directory.SetCurrentDirectory dir
    Net.IniFile.Write "rel.ini", "S", "k", "v"
    AssertTrue Net.File.Exists(dir & "\rel.ini"), "relative path must resolve against the current directory, not the Windows folder"
    Net.Directory.SetCurrentDirectory saved
End Sub
```

- [ ] **Step 2: Run to verify they fail**

Run: `pwsh -File tests\run-tests.ps1 -Office Excel -SkipAddIn`
Expected: `FAIL Test_IniFile_*`.

- [ ] **Step 3: Implement `src/DotNetIniFile.cls` and the facade property**

Every member first resolves `path` with `NewFso().GetAbsolutePathName` (the Win32 API treats a bare name as living in the Windows directory) and raises `DotNetArgument` on empty path; `Write` raises `DotNetDirectoryNotFound` when the parent folder is missing (the API fails silently otherwise). Buffers are `String$(32767, vbNullChar)` passed with `StrPtr`; result length is the return value; lists are split on `vbNullChar` after trimming the double terminator; `DeleteKey` passes `0` for `lpString`, `DeleteSection` passes `0` for `lpKeyName`. `Read` uses a sentinel default (`vbNullChar & "DNV"`) to tell "key exists with empty value" from "key missing", then substitutes `defaultValue` only in the latter case.

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: all `Test_IniFile_*` PASS.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetIniFile.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetIniFile on the Win32 private-profile API"
```

---

### Task 9: DotNetThread and DotNetProcess

**Files:**
- Create: `src/DotNetThread.cls`, `src/DotNetProcess.cls`
- Modify: `src/DotNetVBA.cls` (add `Thread`, `Process` properties)
- Modify: `tests/TestDotNetVBA.bas`, create `tests/MANUAL.md`

**Interfaces:**
- Produces: `Net.Thread.Sleep(ByVal milliseconds As Long)`; `Net.Process.Start(ByVal fileName As String, Optional ByVal arguments As String = "", Optional ByVal windowStyle As Long = 1)`.

- [ ] **Step 1: Write the failing test**

```vb
Public Sub Test_Thread_Sleep()
    Dim t0 As Single: t0 = Timer
    Net.Thread.Sleep 300
    AssertTrue Timer - t0 >= 0.25, "slept " & (Timer - t0)
    On Error Resume Next
    Net.Thread.Sleep -1
    AssertRaised vbObjectError + 1002
    On Error GoTo 0
End Sub

Public Sub Test_Process_StartValidation()
    On Error Resume Next
    Net.Process.Start ""
    AssertRaised vbObjectError + 1001
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: Run to verify they fail, then implement**

`DotNetThread`: `Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)`; negative → `DotNetArgumentOutOfRange`. `DotNetProcess.Start`: empty `fileName` → `DotNetArgument`; `CreateObject("WScript.Shell").Run """" & fileName & """" & IIf(Len(arguments) > 0, " " & arguments, ""), windowStyle, False`.

- [ ] **Step 3: Write `tests/MANUAL.md`**

Checklist with the exact Immediate-window lines to run in Excel, Word and (if available) Access, and the expected observation:
- `Net.Process.Start "notepad.exe"` → Notepad opens.
- `Net.Process.Start Net.Environment.GetFolderPath(SpecialFolder_Desktop)` → Explorer opens the Desktop.
- `?Net.Dialog.OpenFile("Pick", "Text", "*.txt")` → dialog with a Text filter; cancel returns `""`.
- `?UBound(Net.Dialog.OpenFiles("Pick many"))` → multi-select; two files → `1`.
- `?Net.Dialog.SaveFile("Save")` → SaveAs dialog in Excel/Word; `NotSupportedException` in Access.
- `?Net.Dialog.SelectFolder("Folder")` → folder picker.
- A line stating Access is untested by the maintainers and asking testers to report results.

- [ ] **Step 4: Run to verify they pass in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: `Test_Thread_Sleep` and `Test_Process_StartValidation` PASS.

- [ ] **Step 5: Commit**

```bash
git add src/DotNetThread.cls src/DotNetProcess.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas tests/MANUAL.md
git commit -m "feat: add DotNetThread.Sleep and DotNetProcess.Start"
```

---

### Task 10: DotNetDialog

**Files:**
- Create: `src/DotNetDialog.cls`
- Modify: `src/DotNetVBA.cls` (add `Dialog` property)
- Modify: `tests/TestDotNetVBA.bas`

**Interfaces:**
- Produces: `Net.Dialog` with `OpenFile(Optional ByVal title As String = "", Optional ByVal filterName As String = "", Optional ByVal filterPattern As String = "", Optional ByVal initialDirectory As String = "") As String`, `OpenFiles(...same...) As String()`, `SaveFile(...same...) As String`, `SelectFolder(Optional ByVal title As String = "", Optional ByVal initialDirectory As String = "") As String`.
- Constants inlined: `MSO_FILE_DIALOG_SAVE_AS = 2`, `MSO_FILE_DIALOG_FILE_PICKER = 3`, `MSO_FILE_DIALOG_FOLDER_PICKER = 4`.

- [ ] **Step 1: Write the compile-level test**

The dialogs are interactive, so the automated test only proves the class compiles, is reachable, and validates arguments:

```vb
Public Sub Test_Dialog_Reachable()
    AssertTrue Not Net.Dialog Is Nothing
    AssertEqual "DotNetDialog", TypeName(Net.Dialog)
    On Error Resume Next
    Net.Dialog.OpenFile "t", "Text"                                  ' filterName without pattern
    AssertRaised vbObjectError + 1001
    Net.Dialog.OpenFile "t", "", "", "C:\definitely\missing\dir"
    AssertRaised vbObjectError + 1004
    On Error GoTo 0
End Sub
```

- [ ] **Step 2: Run to verify it fails, then implement `src/DotNetDialog.cls`**

One private `ShowDialog(ByVal dialogType As Long, ByVal title As String, ByVal filterName As String, ByVal filterPattern As String, ByVal initialDirectory As String, ByVal multiSelect As Boolean) As Object` that returns the `FileDialog` after `.Show` (or `Nothing` when cancelled). Argument validation happens before any `Application.FileDialog` call so the test never opens a window: `filterName` and `filterPattern` must be both empty or both non-empty (`DotNetArgument`); `initialDirectory` non-empty must exist (`DotNetDirectoryNotFound`). `SaveFile` raises `DotNetNotSupported` when `Application.Name = "Microsoft Access"`. `.InitialFileName = initialDirectory & "\"`; `.Filters.Clear` then `.Filters.Add filterName, filterPattern`; `.AllowMultiSelect = multiSelect`; `.Show = -1` means OK; cancelled → `""` / `EmptyStringArray()`.

- [ ] **Step 3: Run to verify it passes in both hosts**

Run: `pwsh -File tests\run-tests.ps1 -SkipAddIn`
Expected: `PASS Test_Dialog_Reachable` in Excel and Word, no window appears, `FAILED 0`.

- [ ] **Step 4: Commit**

```bash
git add src/DotNetDialog.cls src/DotNetVBA.cls tests/TestDotNetVBA.bas
git commit -m "feat: add DotNetDialog over Office FileDialog"
```

---

### Task 11: Add-in build and add-in boundary test

**Files:**
- Create: `build/build-xlam.ps1`, `tests/TestAddInConsumer.bas`
- Modify: `tests/run-tests.ps1` (replace the ADDIN placeholder)

**Interfaces:**
- Produces: `build\build-xlam.ps1 [-OutputPath <path>]` (default `dist\DotNetVBA.xlam`), exit 0 and prints `BUILT <path> (DotNetVBA <version>)`.
- Produces: `RunAddInConsumerTest(outputPath As String) As Long` in `TestAddInConsumer.bas`, same log format as `RunAllTests`.

- [ ] **Step 1: Write `tests/TestAddInConsumer.bas`**

Standalone module (it cannot use the harness, which is not in the referencing workbook). Contains its own two-line assert and this macro:

```vb
Public Function RunAddInConsumerTest(outputPath As String) As Long
    ' Each check appends PASS/FAIL to outputPath; returns failure count.
    ' 1. Net.Version = "1.0.0"
    ' 2. Net.File.Exists(Net.Path.GetTempPath()) = False and Net.Directory.Exists(Net.Path.GetTempPath()) = True
    ' 3. Dim l As DotNetList: Set l = Net.NewList(): l.Add "a": l.Add "b" -> l(1) = "b" (default member across the boundary)
    ' 4. For Each v In l -> concatenation = "ab" (NewEnum across the boundary)
    ' 5. Net.Environment.GetFolderPath(SpecialFolder_Desktop) is non-empty (enum visible across the boundary)
    ' 6. On Error Resume Next: Net.File.ReadAllText "" -> Err.Number = vbObjectError + 1001 (error numbers survive the boundary)
End Function
```

- [ ] **Step 2: Write `build/build-xlam.ps1`**

Reads the version with `Select-String 'VERSION_STRING As String = "([^"]+)"' src\DotNetVBA.cls`; starts Excel (invisible, `DisplayAlerts = $false`); `Workbooks.Add()`; `VBProject.Name = 'DotNetVBA'`; imports `src\*.bas` then `src\*.cls`; sets `BuiltinDocumentProperties('Title') = 'DotNetVBA'` and `('Comments') = "DotNetVBA $version - .NET-style API for VBA"`; removes an existing output file; `SaveAs($fullPath, 55)` (`xlOpenXMLAddIn`); closes; quits; releases COM; prints the BUILT line. Creates `dist\` if missing.

- [ ] **Step 3: Replace the ADDIN placeholder in `tests/run-tests.ps1`**

Unless `-SkipAddIn` or `-Office Word`: run `build\build-xlam.ps1 -OutputPath (Join-Path $env:TEMP 'dotnetvba-test\DotNetVBA.xlam')`; start Excel; `Workbooks.Open($xlam)`; `$wb = Workbooks.Add()`; `$wb.VBProject.References.AddFromFile($xlam)`; import `tests\TestAddInConsumer.bas`; `$failures = $app.Run('RunAddInConsumerTest', $log)`; print log; close `$wb` without saving; close the add-in workbook; quit; add to the failure total; print `ADDIN: <n> failures`.

- [ ] **Step 4: Run the full suite**

Run: `pwsh -File tests\run-tests.ps1`
Expected: `LINT OK`, Excel and Word import-mode logs with `FAILED 0`, `BUILT ...\DotNetVBA.xlam (DotNetVBA 1.0.0)`, add-in log with six `PASS` lines, `ADDIN: 0 failures`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add build/build-xlam.ps1 tests/TestAddInConsumer.bas tests/run-tests.ps1
git commit -m "build: add xlam build script and add-in boundary test"
```

---

### Task 12: README and v1.0.0

**Files:**
- Create: `README.md`
- Modify: `docs/superpowers/specs/2026-10-27-dotnetvba-design.md` (status line → `Status: implemented`)

- [ ] **Step 1: Write `README.md`** with the sections from spec §9, in this order: title + one-paragraph pitch; the usage example from spec §1; Requirements (64-bit Office 2016+; Excel and Word tested automatically; **Access: supported by design, untested**); Installation (A: `.xlam` from Releases + Tools > References "DotNetVBA"; B: import `src\*.bas` and `src\*.cls`); API reference: one table per class listing members only, copied from spec §6 signatures; Encoding notes (UTF-8 default without BOM; charset names are ADODB names; INI writes ANSI unless the file is already UTF-16); Running the tests (`pwsh -File tests\run-tests.ps1`, trust-access requirement, `tests\MANUAL.md`); Building the add-in; License; a Japanese section `## 日本語` with 概要 / インストール / 使い方 (translation of the pitch, both install options and the usage example). No credits or history section: the README must not mention any earlier library.

The README may contain non-ASCII (it is outside `src/` and `tests/`); save it as UTF-8 without BOM.

- [ ] **Step 2: Verify every public member in `src/` appears in the README tables**

Run (PowerShell): `Select-String -Path src\*.cls,src\*.bas -Pattern '^Public (Function|Sub|Property Get) (\w+)' | ForEach-Object { $_.Matches[0].Groups[2].Value } | Sort-Object -Unique | Where-Object { -not (Select-String -Path README.md -Pattern ('`' + $_) -Quiet) }`
Expected: no output (except `NewEnum`, which is hidden and may be omitted).

- [ ] **Step 3: Full run, then tag**

Run: `pwsh -File tests\run-tests.ps1`
Expected: exit code 0.

```bash
git add README.md docs/superpowers/specs/2026-10-27-dotnetvba-design.md
git commit -m "docs: add README and mark design as implemented"
git tag -a v1.0.0 -m "DotNetVBA 1.0.0"
```

Do **not** push. Report the tag and the path of the built `.xlam` to the owner.
