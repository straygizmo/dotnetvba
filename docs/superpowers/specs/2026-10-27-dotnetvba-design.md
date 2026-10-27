# DotNetVBA — Design

Date: 2026-10-27
Status: implemented

## 1. Intent

DotNetVBA is a set of VBA class modules that let VBA code be written with
.NET API names. It wraps `Scripting.FileSystemObject`, `WScript.Shell`,
`ADODB.Stream`, the Office `FileDialog` and a few Win32 calls behind one
facade object.

The owner's day-to-day pain is file / folder handling and the lack of a
`List`. Those are the center of this project. The library is to be
published on GitHub under the owner's account.

The owner will package the classes in an Excel add-in (`.xlam`) and
reference it from other workbooks. Referenced VBA projects cannot `New`
a class from another project, so the entry point is a public function
`Net()` in a standard module that returns the facade singleton.

Usage shape the owner wants:

```vb
If Net.File.Exists(path) Then
    Dim lines() As String
    lines = Net.File.ReadAllLines(path)
End If

Dim names As DotNetList
Set names = Net.NewList()
names.Add "a"
names.Add "b"
Debug.Print names.Count, names.Contains("a"), names(0)

For Each n In names
    Debug.Print n
Next
```

Inside the same project (classes imported directly) `Dim d As New
DotNet` also works; `Net()` is the documented way.

## 2. Goals

- One entry point: `Net` (a public function returning the `DotNet`
  facade singleton). Static-style .NET classes hang off it as properties
  (`Net.File`, `Net.Directory`, `Net.Path`, `Net.Environment`,
  `Net.Dialog`, `Net.IniFile`, `Net.Thread`, `Net.Process`).
- Usable two ways: (a) import `src\*` into any VBA project, (b) build
  `DotNetVBA.xlam` and reference it from other Excel workbooks.
- Method and property names follow .NET (`System.IO.File`,
  `System.IO.Directory`, `System.IO.Path`, `System.Environment`,
  `System.Collections.Generic.List<T>`, `System.Threading.Thread`,
  `System.Diagnostics.Process`) wherever a sensible VBA equivalent exists.
- A `List` class with .NET semantics: 0-based index, `Count`, `Add`,
  `Insert`, `Remove`, `Contains`, `IndexOf`, `Sort`, `ToArray`, default
  member indexer, and `For Each` support.
- Works unchanged in Excel, Word and Access (and any other Office host
  that exposes `Application.FileDialog`).
- 64-bit Office only. All `Declare` statements use `PtrSafe`.
- No project references required: all COM objects are late-bound
  (`CreateObject`, `As Object`), Office constants are inlined.
- Source is pure ASCII (English comments, no BOM, CRLF line endings) so
  the same bytes import correctly on any system code page.
- Automated tests that run from PowerShell against real Office hosts.
- MIT license. README in English with a short Japanese section.

## 3. Non-goals

- Porting all of .NET. Only the namespaces above, and only members that
  map cleanly to VBA.
- 32-bit Office support.
- Streams (`FileStream`, `StreamReader`), async APIs, LINQ, generics
  with type constraints, delegates / predicates (`Find`, `Exists`,
  `RemoveAll`, `Sort(comparison)`).
- `Dictionary`, `Queue`, `Stack`, `HashSet`. Candidates for a later
  version, not this one.
- Rubberduck or any external test framework dependency.

## 4. Anti-goals (would make this a failure even if it "works")

- A member that has a .NET name but different semantics (e.g. `List`
  that is 1-based, `GetExtension` without the dot, `Directory.Delete`
  that silently deletes a non-empty folder without `recursive:=True`).
- Any Excel-only call (`Application.GetOpenFilename`, `Range`, `Cells`).
- Non-ASCII bytes anywhere in `src/` or `tests/`.
- A green test run that didn't actually touch the file system.

## 5. Constraints

- Hosts: Excel, Word, Access, 64-bit, Office 2016 or later / Microsoft
  365. The test runner verifies Excel and Word. Access is not installed on
  the build machine: README and MANUAL.md must state plainly that Access
  is "supported by design, untested" (owner's decision).
- Language: VBA 7.1. `Option Explicit` in every module. `Declare PtrSafe`
  only. Follow the `vba-coding-conventions` skill when writing code.
- Class naming: every class module is prefixed `DotNet` to avoid
  colliding with host / FSO type names (`File`, `Folder`, `Path`).
  Facade: `DotNet` (not `DotNetVBA`: VBA forbids a module name equal to
  the project name). Classes: `DotNetFile`, `DotNetDirectory`,
  `DotNetPath`, `DotNetEnvironment`, `DotNetList`, `DotNetDialog`,
  `DotNetIniFile`, `DotNetThread`, `DotNetProcess`. Standard module:
  `DotNetVBAGlobals` (holds `Net()`).
- Every class module is `PublicNotCreatable` (`Attribute VB_Exposed =
  True`, `Attribute VB_Creatable = False`) so that a project referencing
  the add-in can declare variables of these types, use `For Each`, the
  default indexer and the `SpecialFolder` enum. Only `Net()` and
  `NewList()` create instances.
- The VBA project name (both when imported into the add-in and in the
  built `.xlam`) is `DotNetVBA`; consumers tick "DotNetVBA" in
  Tools > References.
- Text encoding default is UTF-8 **without BOM** (matches .NET
  `File.WriteAllText` default). Other encodings are passed as an
  ADODB.Stream charset name (`"shift_jis"`, `"utf-16"`, `"utf-8"`).
- Errors are raised with `Err.Raise`, number `vbObjectError + n`, source
  `"DotNetVBA.<Class>.<Member>"`, description prefixed with the .NET
  exception name, e.g.
  `FileNotFoundException: Could not find file 'C:\x.txt'.`
  Fixed numbers:
  - 1001 `ArgumentException` / `ArgumentNullException`
  - 1002 `ArgumentOutOfRangeException`
  - 1003 `FileNotFoundException`
  - 1004 `DirectoryNotFoundException`
  - 1005 `IOException` (exists / not empty / in use)
  - 1006 `InvalidOperationException`
  - 1007 `NotSupportedException`
- The repo is not yet a git repository. The plan starts with `git init`
  and ends with a tagged `v1.0.0` ready to push. Pushing to GitHub is
  done by the owner (or by the builder only if the owner asks).
- Git identity: the machine's global config (`straygizmo
  <straygizmo@gmail.com>`), confirmed by the owner. README links assume
  the repo will live at `github.com/straygizmo/dotnetvba` (my assumption
  from the git identity; trivial to change before pushing).
- "Trust access to the VBA project object model" is ON for both Excel
  and Word on the build machine (owner enabled Word manually).

## 6. Public API

Return type notes: "String()" means a 0-based `String` array. When there
are no results, return an empty array with `UBound = -1` (obtainable via
`Split(vbNullString)`), never `Nothing` or an uninitialized array.

### 6.1 DotNet (facade)

| Member | Returns | Notes |
|---|---|---|
| `File` | `DotNetFile` | lazily created singleton |
| `Directory` | `DotNetDirectory` | " |
| `Path` | `DotNetPath` | " |
| `Environment` | `DotNetEnvironment` | " |
| `Dialog` | `DotNetDialog` | " |
| `IniFile` | `DotNetIniFile` | " |
| `Thread` | `DotNetThread` | " |
| `Process` | `DotNetProcess` | " |
| `NewList()` | `DotNetList` | new instance every call |
| `Version` | `String` | `"1.0.0"` |

### 6.1a DotNetVBAGlobals (standard module)

| Member | Signature |
|---|---|
| `Net` | `() As DotNet` — returns a module-level singleton, created on first call. Public, so it is callable from referencing projects as `Net.File.Exists(...)`. |

### 6.2 DotNetFile (System.IO.File)

| Member | Signature |
|---|---|
| `Exists` | `(path As String) As Boolean` — False for folders, like .NET |
| `Copy` | `(source As String, dest As String, Optional overwrite As Boolean = False)` — IOException if dest exists and not overwrite |
| `Move` | `(source As String, dest As String, Optional overwrite As Boolean = False)` |
| `Delete` | `(path As String)` — no error if missing (like .NET) |
| `ReadAllText` | `(path As String, Optional encoding As String = "utf-8") As String` |
| `ReadAllLines` | `(path As String, Optional encoding As String = "utf-8") As String()` — accepts CRLF, LF and CR; a trailing newline does not produce an extra empty element |
| `WriteAllText` | `(path As String, contents As String, Optional encoding As String = "utf-8")` — overwrites; UTF-8 written without BOM |
| `WriteAllLines` | `(path As String, lines As Variant, Optional encoding As String = "utf-8")` — `lines` is a String array or a `DotNetList`; each line terminated by CRLF |
| `AppendAllText` | `(path As String, contents As String, Optional encoding As String = "utf-8")` — creates file if missing |
| `AppendAllLines` | `(path As String, lines As Variant, Optional encoding As String = "utf-8")` |
| `GetCreationTime` | `(path As String) As Date` |
| `GetLastWriteTime` | `(path As String) As Date` |
| `GetLastAccessTime` | `(path As String) As Date` |
| `GetLength` | `(path As String) As Double` — not in .NET's `File`, kept because FileInfo.Length is the common need; name chosen by the builder's design authority: `GetLength` |
| `GetAttributes` | `(path As String) As Long` — FSO attribute bits |

Reading a missing file raises `FileNotFoundException`. Writing into a
missing folder raises `DirectoryNotFoundException`.

Implementation note (left to builder, but required behavior): ADODB.Stream
writes a BOM for `utf-8`; strip it so output matches .NET. Reading must
tolerate a BOM if present.

### 6.3 DotNetDirectory (System.IO.Directory)

| Member | Signature |
|---|---|
| `Exists` | `(path As String) As Boolean` |
| `CreateDirectory` | `(path As String)` — creates all missing parents; no error if it already exists |
| `Delete` | `(path As String, Optional recursive As Boolean = False)` — IOException if non-empty and not recursive; DirectoryNotFoundException if missing |
| `Move` | `(source As String, dest As String)` |
| `GetFiles` | `(path As String, Optional searchPattern As String = "*", Optional recursive As Boolean = False) As String()` — full paths |
| `GetDirectories` | `(path As String, Optional searchPattern As String = "*", Optional recursive As Boolean = False) As String()` — full paths |
| `GetParent` | `(path As String) As String` — `""` for a root |
| `GetCurrentDirectory` | `() As String` |
| `SetCurrentDirectory` | `(path As String)` |
| `GetCreationTime` / `GetLastWriteTime` | `(path As String) As Date` |

`searchPattern` supports `*` and `?` only, matched case-insensitively
against the file / folder name. `recursive:=True` corresponds to .NET's
`SearchOption.AllDirectories`; a Boolean is used instead of an enum to
keep call sites short.

### 6.4 DotNetPath (System.IO.Path)

| Member | Signature |
|---|---|
| `Combine` | `(ParamArray parts() As Variant) As String` — inserts `\` between parts exactly once; a rooted later part discards earlier parts (as .NET) |
| `GetFileName` | `(path As String) As String` |
| `GetFileNameWithoutExtension` | `(path As String) As String` |
| `GetExtension` | `(path As String) As String` — includes the leading dot; `""` if none |
| `GetDirectoryName` | `(path As String) As String` — no trailing separator |
| `ChangeExtension` | `(path As String, extension As String) As String` — `extension` with or without dot; empty removes it |
| `HasExtension` | `(path As String) As Boolean` |
| `IsPathRooted` | `(path As String) As Boolean` |
| `GetPathRoot` | `(path As String) As String` |
| `GetFullPath` | `(path As String) As String` — resolves relative to current directory, normalizes `.`/`..` |
| `GetTempPath` | `() As String` — with trailing `\` (as .NET) |
| `GetTempFileName` | `() As String` — creates a zero-byte file and returns its path |
| `GetRandomFileName` | `() As String` |
| `DirectorySeparatorChar` | `Property Get As String` — `"\"` |

Pure string functions; they do not touch the file system except
`GetFullPath`, `GetTempPath`, `GetTempFileName`.

### 6.5 DotNetEnvironment (System.Environment)

```vb
Public Enum SpecialFolder
    SpecialFolder_Desktop
    SpecialFolder_MyDocuments
    SpecialFolder_ApplicationData
    SpecialFolder_LocalApplicationData
    SpecialFolder_UserProfile
    SpecialFolder_ProgramFiles
    SpecialFolder_Startup
    SpecialFolder_StartMenu
    SpecialFolder_Templates
    SpecialFolder_Favorites
    SpecialFolder_Recent
    SpecialFolder_SendTo
End Enum
```

| Member | Signature |
|---|---|
| `GetFolderPath` | `(folder As SpecialFolder) As String` |
| `NewLine` | `Property Get As String` — `vbCrLf` |
| `UserName` | `Property Get As String` |
| `MachineName` | `Property Get As String` |
| `CurrentDirectory` | `Property Get / Let As String` |
| `GetEnvironmentVariable` | `(name As String) As String` — `""` if missing |
| `ExpandEnvironmentVariables` | `(text As String) As String` |
| `Is64BitProcess` | `Property Get As Boolean` — always True in this library, provided for symmetry |

Enum members are prefixed `SpecialFolder_` because VBA enums share one
global namespace per project. The `SpecialFolder` enum is declared
`Public` inside the `DotNetEnvironment` class module.

### 6.6 DotNetList (System.Collections.Generic.List<T>)

Items are `Variant`; values and objects can both be stored (objects are
assigned with `Set` internally, detected by `IsObject`). Index is
0-based. `Item` is the default member (`VB_UserMemId = 0`) and
`NewEnum` enables `For Each` (`VB_UserMemId = -4`, hidden).

| Member | Signature |
|---|---|
| `Count` | `Property Get As Long` |
| `Item` | `Property Get (index As Long) As Variant` (default) / `Property Let` / `Property Set` |
| `Add` | `(item As Variant)` |
| `AddRange` | `(items As Variant)` — array or `DotNetList` |
| `Insert` | `(index As Long, item As Variant)` — `index = Count` appends |
| `Remove` | `(item As Variant) As Boolean` — first match |
| `RemoveAt` | `(index As Long)` |
| `RemoveRange` | `(index As Long, count As Long)` |
| `Clear` | `()` |
| `Contains` | `(item As Variant) As Boolean` |
| `IndexOf` | `(item As Variant) As Long` — `-1` if absent |
| `LastIndexOf` | `(item As Variant) As Long` |
| `GetRange` | `(index As Long, count As Long) As DotNetList` |
| `Sort` | `()` — values only; strings compared with `vbBinaryCompare`; raises `InvalidOperationException` if any item is an object |
| `Reverse` | `()` |
| `ToArray` | `() As Variant()` — 0-based |
| `Join` | `(Optional separator As String = ",") As String` — not in .NET's List (it is `String.Join`); included because it is the overwhelmingly common VBA need |
| `NewEnum` | hidden, `IUnknown`, for `For Each` |

Equality for `Contains` / `IndexOf` / `Remove`: values use `=`;
objects use `Is`; a value and an object are never equal.
Out-of-range index raises `ArgumentOutOfRangeException`.
Backing store is a `Variant` array that doubles its capacity when full.

### 6.7 DotNetDialog (no .NET static equivalent; Office FileDialog)

All members use `Application.FileDialog(n)` with inlined constants
(`msoFileDialogFilePicker = 3`, `msoFileDialogFolderPicker = 4`,
`msoFileDialogSaveAs = 2`). `Application` resolves to the host's
Application object in Excel, Word and Access.

| Member | Signature |
|---|---|
| `OpenFile` | `(Optional title As String, Optional filterName As String, Optional filterPattern As String, Optional initialDirectory As String) As String` — `""` if cancelled |
| `OpenFiles` | same args `As String()` — empty array if cancelled |
| `SaveFile` | same args `As String` — raises `NotSupportedException` in Access (Access has no SaveAs FileDialog) |
| `SelectFolder` | `(Optional title As String, Optional initialDirectory As String) As String` |

Not covered by automated tests (interactive). A manual checklist lives
in `tests/MANUAL.md`.

### 6.8 DotNetIniFile (no .NET equivalent; Win32 private profile API)

Replaces the hand-written INI parser with `GetPrivateProfileStringW`,
`WritePrivateProfileStringW` and `GetPrivateProfileSectionNamesW`
(`Declare PtrSafe`). Behavior therefore matches every other Windows
program that uses INI files.

| Member | Signature |
|---|---|
| `ReadString` | `(path As String, section As String, keyName As String, Optional defaultValue As String = "") As String` |
| `WriteString` | `(path As String, section As String, keyName As String, value As String)` — creates file / section as needed |
| `DeleteKey` | `(path As String, section As String, keyName As String)` |
| `DeleteSection` | `(path As String, section As String)` |
| `GetSections` | `(path As String) As String()` |
| `GetKeys` | `(path As String, section As String) As String()` |

(`Write` is a VBA keyword and cannot name a procedure, so the pair is
`ReadString` / `WriteString`, mirroring the Win32 API names.)

Parameter is `keyName`, not `key`, because `Key` is reserved-ish in VBA
(see the `vba-coding-conventions` skill).

Known limitation (document in README): the Win32 API writes ANSI unless
the file already exists as UTF-16 with BOM.

### 6.9 DotNetThread (System.Threading.Thread)

| Member | Signature |
|---|---|
| `Sleep` | `(milliseconds As Long)` — kernel32 `Sleep`, `PtrSafe` |

### 6.10 DotNetProcess (System.Diagnostics.Process)

| Member | Signature |
|---|---|
| `Start` | `(fileName As String, Optional arguments As String = "", Optional windowStyle As Long = 1)` — `WScript.Shell.Run` with quoting; opens documents with their associated application, like .NET's `UseShellExecute = True` |

## 7. Repository layout

```
dotnetvba/
  README.md                 English; short Japanese section at the end
  LICENSE                   MIT, copyright holder "straygizmo" (owner's git identity)
  .gitignore                Office temp files (~$*), *.xlsm test scratch
  .gitattributes            *.cls *.bas text eol=crlf
  src/
    DotNetVBAGlobals.bas
    DotNetVBAInternal.bas   Option Private Module: shared helpers, error enum
    DotNet.cls
    DotNetFile.cls
    DotNetDirectory.cls
    DotNetPath.cls
    DotNetEnvironment.cls
    DotNetList.cls
    DotNetDialog.cls
    DotNetIniFile.cls
    DotNetThread.cls
    DotNetProcess.cls
  tests/
    TestDotNetVBA.bas       test module (own tiny assert helpers)
    TestAddInConsumer.bas   tiny consumer macro used by the add-in test
    run-tests.ps1           PowerShell runner (see 8)
    MANUAL.md               checklist for Dialog
  build/
    build-xlam.ps1          builds dist/DotNetVBA.xlam from src/ (see 7.1)
  dist/                     gitignored; release artifact only
  docs/superpowers/specs/   this document
```

The five pre-existing `.cls` files in the working folder are deleted
before the first commit and are not mentioned anywhere in the published
repository (owner's decision).

### 7.1 Add-in build (`build/build-xlam.ps1`)

Starts Excel via COM, creates a workbook, sets `VBProject.Name =
"DotNetVBA"`, imports `src\*.bas` and `src\*.cls`, sets the workbook
`Title`/`Comments` (shown in the Add-ins dialog) to "DotNetVBA x.y.z",
saves as `dist\DotNetVBA.xlam` (`FileFormat = 55`), quits. The `.xlam`
is attached to each GitHub Release, not committed (my call: binary diffs
are noise; the script is the source of truth).

## 8. Testing

### 8.1 Test module (`tests/TestDotNetVBA.bas`)

- Plain `.bas` module, no external framework. Provides `AssertEqual`,
  `AssertTrue`, `AssertFalse`, `AssertRaises(expectedNumber)` pattern via
  `On Error`, and a `RunAllTests(outputPath As String) As Long` entry
  point that runs every `Test_*` procedure, writes one line per test
  (`PASS`/`FAIL` + name + message) to `outputPath`, and returns the
  failure count.
- Each test gets a fresh scratch folder under `Path.GetTempPath` and
  deletes it afterwards.
- Add-in test (Excel only): `run-tests.ps1` builds the `.xlam` via
  `build-xlam.ps1`, opens it, creates a fresh workbook, adds a reference
  to the add-in's VBProject, imports `TestAddInConsumer.bas` (which calls
  `Net.File.Exists`, `Net.NewList`, a `For Each`, `names(0)` and the
  `SpecialFolder` enum), runs it and asserts success. This proves the
  `PublicNotCreatable` exposure and the attributes survive the add-in
  boundary.
- Coverage: every public member of `File`, `Directory`, `Path`,
  `Environment` (`GetFolderPath` for Desktop and MyDocuments at least),
  `List`, `IniFile`, `Thread.Sleep` (elapsed time ≥ requested),
  `Process.Start` is **not** auto-tested (launches a window).
- Encoding tests: write and read back a string containing non-ASCII
  characters (built with `ChrW`, so the source stays ASCII) in `utf-8`,
  `utf-16` and `shift_jis`; assert the UTF-8 file has no BOM by reading
  its first bytes.
- Line-ending tests: `ReadAllLines` on CRLF, LF, CR and mixed input,
  with and without trailing newline.

### 8.2 Runner (`tests/run-tests.ps1`)

```
.\tests\run-tests.ps1 [-Office Excel|Word|All]   (default: All)

(The parameter is `-Office`, not `-Host`: `$Host` is a read-only
PowerShell automatic variable.)
```

For each host: start the application via COM (invisible), create a new
document / workbook, import every `src\*.cls` and `tests\*.bas` through
`VBProject.VBComponents.Import`, call `Application.Run "RunAllTests",
<temp log path>`, read the log, print it, quit the host without saving.
Exit code = total failures. Fails fast with a clear message if "Trust
access to the VBA project object model" is off.

Word note: `Application.Run` in Word takes the macro name as
`"RunAllTests"` with arguments via `Run(MacroName, varg1)`; the builder
verifies the exact overload.

## 9. README contents

- One-paragraph pitch, the usage example from section 1.
- Installation, two options: (1) download `DotNetVBA.xlam` from
  Releases, put it in your add-ins folder, enable it, then in your
  workbook's VBE tick "DotNetVBA" under Tools > References; (2) import
  `src\*.bas` and `src\*.cls` into your own project (any host). No
  other references needed either way.
- Requirements: 64-bit Office, Excel / Word / Access.
- API table per class (members only; one line each).
- Encoding and INI limitations.
- Running the tests.
- Japanese section: 概要・インストール・使い方 (a short translation of
  the pitch and installation).

## 10. Left to the builder

- `NewEnum` builds a snapshot `Collection` from the Variant array on each
  call and returns its `[_NewEnum]`.
- Internal helper layout (shared FSO instance per class vs. per call;
  shared encoding helpers between `File` members).
- Exact technique for BOM-less UTF-8 output and BOM-tolerant input.
- `Path.GetFullPath` normalization algorithm.
- Wildcard-to-`Like` pattern escaping details.
- Error-message wording beyond the exception-name prefix.
- Test fixture naming and the assert helper implementation.
- Whether `run-tests.ps1` also runs Access when installed (optional
  extra; Excel and Word are required).

## 11. Decisions made in conversation (owner's)

- Name `DotNetVBA`, entry point `Net.File.Exists(path)` via a public
  `Net()` function (chosen because `.xlam` consumers cannot `New`).
- Packaged as an Excel add-in (`.xlam`) for daily use.
- The pre-existing `.cls` files are deleted and never referenced.
- Scope focus: files, folders, `List`.
- `List` is 0-based.
- Word / Access / Excel common. 64-bit only.
- English comments, ASCII source.
- Automated tests against real Office via COM; trust access is enabled
  in Excel and Word.
- Git identity stays as the global config. Access documented as untested.

## 12. Decisions made by Claude (owner should glance at these)

- MIT license.
- Class prefix `DotNet` on every module.
- `Directory.GetFiles` returns full paths and uses a `recursive` Boolean.
- `List.Join` and `File.GetLength` added despite not being on the .NET
  class of the same name.
- INI reimplemented on the Win32 private-profile API.
- `.xlam` built by script and attached to Releases, not committed.
- Standard module named `DotNetVBAGlobals`; project name `DotNetVBA`.
- UTF-8 without BOM as the default text encoding.
- Error numbers 1001–1007 as listed in section 5.
