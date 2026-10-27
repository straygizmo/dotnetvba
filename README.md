# DotNetVBA

DotNetVBA is a set of VBA class modules that let you write VBA with .NET API names: `File.Exists`, `Directory.GetFiles`, `Path.Combine`, `List` with 0-based indexing, and more. It wraps `Scripting.FileSystemObject`, `WScript.Shell`, `ADODB.Stream`, the Office `FileDialog` and a few Win32 calls behind one entry point, `Net`. It works in Excel, Word and Access (64-bit), needs no project references, and can be imported as source or used as an Excel add-in (`.xlam`).

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

`Net()` is a public function in a standard module that returns the facade singleton. Inside a project that imported the classes directly, `Dim d As New DotNet` also works.

## Requirements

- 64-bit Office 2016 or later / Microsoft 365 (all `Declare` statements are `PtrSafe`; 32-bit is not supported).
- Excel and Word: tested automatically against real Office.
- Access: supported by design, untested.

## Installation

**A. Add-in.** Download `DotNetVBA.xlam` from [Releases](https://github.com/straygizmo/dotnetvba/releases), put it in your add-ins folder and enable it. In your workbook's VBE, tick "DotNetVBA" under Tools > References.

**B. Source.** Import `src\*.bas` and `src\*.cls` into your own project (any host).

No other references are needed either way.

## API reference

All classes are prefixed `DotNet` and are `PublicNotCreatable`; get them from `Net`. "String()" is a 0-based array; no results gives an empty array (`UBound = -1`).

### Net (module `DotNetVBAGlobals`) and DotNet (facade)

| Member | Description |
|---|---|
| `Net() As DotNet` | Facade singleton |
| `File` | `DotNetFile` |
| `Directory` | `DotNetDirectory` |
| `Path` | `DotNetPath` |
| `Environment` | `DotNetEnvironment` |
| `Dialog` | `DotNetDialog` |
| `IniFile` | `DotNetIniFile` |
| `Thread` | `DotNetThread` |
| `Process` | `DotNetProcess` |
| `NewList() As DotNetList` | New list on every call |
| `Version As String` | `"1.0.0"` |

### File (System.IO.File)

| Member | Description |
|---|---|
| `Exists(path) As Boolean` | False for folders |
| `Copy(source, dest, [overwrite = False])` | IOException if dest exists |
| `Move(source, dest, [overwrite = False])` | |
| `Delete(path)` | No error if missing |
| `ReadAllText(path, [encoding = "utf-8"]) As String` | |
| `ReadAllLines(path, [encoding = "utf-8"]) As String()` | CRLF, LF or CR; no extra trailing element |
| `WriteAllText(path, contents, [encoding = "utf-8"])` | Overwrites |
| `WriteAllLines(path, lines As Variant, [encoding = "utf-8"])` | `lines`: String array or `DotNetList` |
| `AppendAllText(path, contents, [encoding = "utf-8"])` | Creates the file if missing |
| `AppendAllLines(path, lines As Variant, [encoding = "utf-8"])` | |
| `GetCreationTime(path) As Date` | |
| `GetLastWriteTime(path) As Date` | |
| `GetLastAccessTime(path) As Date` | |
| `GetLength(path) As Double` | Size in bytes |
| `GetAttributes(path) As Long` | FSO attribute bits |

### Directory (System.IO.Directory)

| Member | Description |
|---|---|
| `Exists(path) As Boolean` | |
| `CreateDirectory(path)` | Creates missing parents |
| `Delete(path, [recursive = False])` | IOException if not empty and not recursive |
| `Move(sourceDirName, destDirName)` | |
| `GetFiles(path, [searchPattern = "*"], [recursive = False]) As String()` | Full paths |
| `GetDirectories(path, [searchPattern = "*"], [recursive = False]) As String()` | Full paths |
| `GetParent(path) As String` | `""` for a root |
| `GetCurrentDirectory() As String` | |
| `SetCurrentDirectory(path)` | |
| `GetCreationTime(path) As Date` | |
| `GetLastWriteTime(path) As Date` | |

`searchPattern` supports `*` and `?`, case-insensitive.

### Path (System.IO.Path)

| Member | Description |
|---|---|
| `Combine(ParamArray parts) As String` | A rooted later part discards earlier ones |
| `GetFileName(path) As String` | |
| `GetFileNameWithoutExtension(path) As String` | |
| `GetExtension(path) As String` | Includes the dot |
| `GetDirectoryName(path) As String` | |
| `ChangeExtension(path, extension) As String` | |
| `HasExtension(path) As Boolean` | |
| `IsPathRooted(path) As Boolean` | |
| `GetPathRoot(path) As String` | |
| `GetFullPath(path) As String` | Relative to the current directory |
| `GetTempPath() As String` | Trailing `\` |
| `GetTempFileName() As String` | Creates a zero-byte file |
| `GetRandomFileName() As String` | |
| `DirectorySeparatorChar As String` | `"\"` |

### Environment (System.Environment)

| Member | Description |
|---|---|
| `GetFolderPath(folder As SpecialFolder) As String` | `SpecialFolder_Desktop`, `_MyDocuments`, `_ApplicationData`, `_LocalApplicationData`, `_UserProfile`, `_ProgramFiles`, `_Startup`, `_StartMenu`, `_Templates`, `_Favorites`, `_Recent`, `_SendTo` |
| `NewLine As String` | |
| `UserName As String` | |
| `MachineName As String` | |
| `CurrentDirectory As String` | Get / Let |
| `GetEnvironmentVariable(variableName) As String` | `""` if missing |
| `ExpandEnvironmentVariables(text) As String` | |
| `Is64BitProcess As Boolean` | Always True |

### List (System.Collections.Generic.List&lt;T&gt;)

Items are `Variant` (values or objects). Index is 0-based; `Item` is the default member; `For Each` is supported.

| Member | Description |
|---|---|
| `Count As Long` | |
| `Item(index As Long) As Variant` | Default; Get / Let / Set |
| `Add(value)` | |
| `AddRange(items As Variant)` | Array or `DotNetList` |
| `Insert(index, value)` | `index = Count` appends |
| `Remove(value) As Boolean` | First match |
| `RemoveAt(index)` | |
| `RemoveRange(index, itemCount)` | |
| `Clear()` | |
| `Contains(value) As Boolean` | |
| `IndexOf(value) As Long` | -1 if absent |
| `LastIndexOf(value) As Long` | |
| `GetRange(index, itemCount) As DotNetList` | |
| `Sort()` | Values only (binary compare); InvalidOperationException on objects |
| `Reverse()` | |
| `ToArray() As Variant()` | 0-based |
| `Join([separator = ","]) As String` | |

### Dialog (Office FileDialog)

| Member | Description |
|---|---|
| `OpenFile([title], [filterName], [filterPattern], [initialDirectory]) As String` | `""` if cancelled |
| `OpenFiles([title], [filterName], [filterPattern], [initialDirectory]) As String()` | Empty array if cancelled |
| `SaveFile([title], [filterName], [filterPattern], [initialDirectory]) As String` | NotSupportedException in Access |
| `SelectFolder([title], [initialDirectory]) As String` | |

### IniFile (Win32 private profile API)

| Member | Description |
|---|---|
| `ReadString(path, section, keyName, [defaultValue = ""]) As String` | |
| `WriteString(path, section, keyName, value)` | Creates file / section as needed |
| `DeleteKey(path, section, keyName)` | |
| `DeleteSection(path, section)` | |
| `GetSections(path) As String()` | |
| `GetKeys(path, section) As String()` | |

### Thread (System.Threading.Thread) and Process (System.Diagnostics.Process)

| Member | Description |
|---|---|
| `Sleep(milliseconds)` | Thread |
| `Start(fileName, [arguments = ""], [windowStyle = 1])` | Process: opens documents with their associated application |

## Notes and limitations

**Encoding**
- Default encoding is UTF-8, written without BOM (like .NET) and read with or without one.
- Other encodings are ADODB.Stream charset names (`utf-8`, `utf-16`, `shift_jis`) and keep ADODB's BOM behavior (`utf-16` writes `FF FE`).
- `AppendAllText` / `AppendAllLines` rewrite the whole file; an existing UTF-8 BOM is dropped.
- `IniFile` writes ANSI unless the file already exists as UTF-16 with BOM (Win32 behavior).

**Behavior**
- `Directory.SetCurrentDirectory` cannot target a UNC path (VBA `ChDir` limitation); it raises IOException.
- `Directory.GetFiles` pattern `*.*` does not match extension-less names; use `*`.
- `List` enumeration takes a snapshot when `For Each` starts, so changing the list inside the loop does not raise (unlike .NET). Do not mutate a list while enumerating it.
- `Dialog.SaveFile` raises NotSupportedException in Access.
- Access is supported by design but untested by the maintainers.

**Errors** are raised with `Err.Number = vbObjectError + n` and an `Err.Description` that starts with the .NET exception name, for example `FileNotFoundException: Could not find file 'C:\x.txt'.`

| n | Exception |
|---|---|
| 1001 | ArgumentException / ArgumentNullException |
| 1002 | ArgumentOutOfRangeException |
| 1003 | FileNotFoundException |
| 1004 | DirectoryNotFoundException |
| 1005 | IOException |
| 1006 | InvalidOperationException |
| 1007 | NotSupportedException |

## Running the tests

```
pwsh -File tests\run-tests.ps1 [-Office Excel|Word|All] [-SkipAddIn]
```

The runner drives real Excel and Word through COM, builds the add-in and tests it through a consumer workbook. It needs "Trust access to the VBA project object model" enabled in each host's Trust Center. The exit code is the number of failures. Interactive `Dialog` checks are listed in `tests\MANUAL.md`.

## Building the add-in

```
pwsh -File build\build-xlam.ps1 [-OutputPath ...]
```

This produces `dist\DotNetVBA.xlam` (VBA project name `DotNetVBA`) from `src\`. Excel must be installed.

## License

MIT. See [LICENSE](LICENSE). Repository: <https://github.com/straygizmo/dotnetvba>

## 日本語

### 概要

DotNetVBA は、.NET の API 名 (`File.Exists`、`Directory.GetFiles`、`Path.Combine`、0 始まりの `List` など) で VBA を書けるようにするクラスモジュール集です。`Scripting.FileSystemObject`、`WScript.Shell`、`ADODB.Stream`、Office の `FileDialog`、一部の Win32 API を、`Net` という 1 つの入口にまとめています。Excel / Word / Access (64 ビット) で動作し、参照設定は不要です。ソースのインポートでも、Excel アドイン (`.xlam`) としても使えます。Excel と Word は実機で自動テスト済み、Access は設計上は対応していますが未テストです。

### インストール

- **A. アドイン**: [Releases](https://github.com/straygizmo/dotnetvba/releases) から `DotNetVBA.xlam` をダウンロードしてアドインフォルダに置き、有効化します。使用するブックの VBE で、ツール > 参照設定 の "DotNetVBA" にチェックを入れます。
- **B. ソース**: `src\*.bas` と `src\*.cls` を自分のプロジェクトにインポートします。

### 使い方

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

既定の文字コードは BOM なし UTF-8 です。他の文字コードは ADODB の名前 (`shift_jis`、`utf-16` など) で指定します。
