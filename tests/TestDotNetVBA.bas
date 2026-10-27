Attribute VB_Name = "TestDotNetVBA"
Option Explicit

' Test module for DotNetVBA.
' Each test is: Public Sub Test_<Class>_<Behavior>()
' RunAllTests discovers them by scanning this module's code.

Private Const MODULE_NAME As String = "TestDotNetVBA"
Private Const DISPATCH_MODULE As String = "TestDotNetVBADispatch"
Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const ASSERT_ERROR As Long = vbObjectError + 9999

' Temp folder of the running test; deleted by RunAllTests after each test.
Private mTempDir As String
' First assertion failure of the running test. Kept because a failure raised
' while the test has On Error Resume Next active would otherwise be lost.
Private mFailMessage As String

' ===== Runner =====

' Runs every Public Sub Test_* of this module, writes one PASS/FAIL line per
' test plus a TOTAL line to outputPath, and returns the failure count.
' A procedure started by Application.Run does not pass its errors back to the
' caller, so the tests are called directly from a generated dispatcher module
' whose RunOneTest traps the error and returns its description.
Public Function RunAllTests(ByVal outputPath As String) As Long
    Dim fso As Object
    Dim logFile As Object
    Dim components As Object
    Dim testNames As Collection
    Dim testName As Variant
    Dim errText As String
    Dim totalCount As Long
    Dim failCount As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Unicode (UTF-16) so any failure text can be written.
    Set logFile = fso.CreateTextFile(outputPath, True, True)
    Set components = FindOwnComponents()
    Set testNames = FindTestNames(components(MODULE_NAME).CodeModule)
    AddDispatcher components, testNames

    For Each testName In testNames
        totalCount = totalCount + 1
        mTempDir = vbNullString
        mFailMessage = vbNullString
        errText = Application.Run(DISPATCH_MODULE & ".RunOneTest", CStr(testName))
        DeleteTempDir fso
        If mFailMessage <> vbNullString Then errText = mFailMessage
        If errText = vbNullString Then
            logFile.WriteLine "PASS " & testName
        Else
            failCount = failCount + 1
            logFile.WriteLine "FAIL " & testName & ": " & errText
        End If
    Next testName

    RemoveDispatcher components
    logFile.WriteLine "TOTAL " & totalCount & " FAILED " & failCount
    logFile.Close
    RunAllTests = failCount
End Function

' VBComponents of the project holding this module (in Word, ActiveVBProject
' can be Normal rather than the test document).
Private Function FindOwnComponents() As Object
    Dim proj As Object
    Dim probe As Object
    For Each proj In Application.VBE.VBProjects
        Set probe = Nothing
        On Error Resume Next
        Set probe = proj.VBComponents(MODULE_NAME)
        On Error GoTo 0
        If Not probe Is Nothing Then
            Set FindOwnComponents = proj.VBComponents
            Exit Function
        End If
    Next proj
    Err.Raise ASSERT_ERROR, MODULE_NAME, "Cannot find the project containing " & MODULE_NAME
End Function

Private Function FindTestNames(ByVal codeMod As Object) As Collection
    Dim names As New Collection
    Dim lineIdx As Long
    Dim lineText As String
    Dim procName As String
    For lineIdx = 1 To codeMod.CountOfLines
        lineText = codeMod.Lines(lineIdx, 1)
        If Left$(lineText, 16) = "Public Sub Test_" Then
            procName = Mid$(lineText, 12)
            procName = Trim$(Left$(procName, InStr(procName, "(") - 1))
            names.Add procName
        End If
    Next lineIdx
    Set FindTestNames = names
End Function

Private Sub AddDispatcher(ByVal components As Object, ByVal testNames As Collection)
    Dim code As String
    Dim testName As Variant
    RemoveDispatcher components
    code = "Option Explicit" & vbCrLf & _
           "Public Function RunOneTest(ByVal procName As String) As String" & vbCrLf & _
           "    On Error GoTo Failed" & vbCrLf & _
           "    Select Case procName" & vbCrLf
    For Each testName In testNames
        code = code & "        Case """ & testName & """: " & MODULE_NAME & "." & testName & vbCrLf
    Next testName
    code = code & _
           "        Case Else: RunOneTest = ""no such test"": Exit Function" & vbCrLf & _
           "    End Select" & vbCrLf & _
           "    Exit Function" & vbCrLf & _
           "Failed:" & vbCrLf & _
           "    RunOneTest = Err.Description" & vbCrLf & _
           "    If RunOneTest = vbNullString Then RunOneTest = ""Error "" & Err.Number" & vbCrLf & _
           "End Function" & vbCrLf
    With components.Add(VBEXT_CT_STDMODULE)
        .Name = DISPATCH_MODULE
        .CodeModule.AddFromString code
    End With
End Sub

Private Sub RemoveDispatcher(ByVal components As Object)
    Dim existing As Object
    On Error Resume Next
    Set existing = components(DISPATCH_MODULE)
    On Error GoTo 0
    If Not existing Is Nothing Then components.Remove existing
End Sub

Private Sub DeleteTempDir(ByVal fso As Object)
    If mTempDir = vbNullString Then Exit Sub
    On Error Resume Next
    If fso.FolderExists(mTempDir) Then fso.DeleteFolder mTempDir, True
    On Error GoTo 0
    mTempDir = vbNullString
End Sub

' ===== Assertions =====

Private Sub Fail(ByVal message As String)
    If mFailMessage = vbNullString Then mFailMessage = message
    Err.Raise ASSERT_ERROR, MODULE_NAME, message
End Sub

Private Function WithMessage(ByVal text As String, ByVal message As String) As String
    If message = vbNullString Then
        WithMessage = text
    Else
        WithMessage = message & " - " & text
    End If
End Function

Private Function Describe(ByRef value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            Describe = "Nothing"
        Else
            Describe = "<" & TypeName(value) & ">"
        End If
    ElseIf IsArray(value) Then
        Describe = "<" & TypeName(value) & ">"
    ElseIf IsNull(value) Then
        Describe = "Null"
    ElseIf IsEmpty(value) Then
        Describe = "Empty"
    ElseIf VarType(value) = vbString Then
        Describe = """" & value & """"
    Else
        Describe = CStr(value)
    End If
End Function

' Compares like AssertEqual. Leaves Err untouched, so a test can assert on
' values between the failing call and AssertRaised.
Private Function ValuesEqual(ByRef expected As Variant, ByRef actual As Variant) As Boolean
    Dim savedNumber As Long
    Dim savedSource As String
    Dim savedDesc As String
    savedNumber = Err.Number
    savedSource = Err.Source
    savedDesc = Err.Description

    On Error GoTo Mismatch
    If IsObject(expected) Or IsObject(actual) Then
        If IsObject(expected) And IsObject(actual) Then ValuesEqual = (expected Is actual)
    ElseIf IsArray(expected) Or IsArray(actual) Then
        ValuesEqual = False
    ElseIf IsNull(expected) Or IsNull(actual) Then
        ValuesEqual = (IsNull(expected) And IsNull(actual))
    ElseIf VarType(expected) = vbString Or VarType(actual) = vbString Then
        ValuesEqual = (CStr(expected) = CStr(actual))
    Else
        ValuesEqual = (expected = actual)
    End If
    GoTo Done
Mismatch:
    ValuesEqual = False
    Resume Done
Done:
    On Error GoTo 0
    If savedNumber <> 0 Then
        Err.Number = savedNumber
        Err.Source = savedSource
        Err.Description = savedDesc
    End If
End Function

Public Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String)
    If Not ValuesEqual(expected, actual) Then
        Fail WithMessage("AssertEqual expected " & Describe(expected) & " but was " & Describe(actual), message)
    End If
End Sub

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String)
    If Not condition Then Fail WithMessage("AssertTrue failed", message)
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String)
    If condition Then Fail WithMessage("AssertFalse failed", message)
End Sub

Public Sub AssertArrayEqual(ByVal expectedArray As Variant, ByVal actualArray As Variant, Optional ByVal message As String)
    Dim idx As Long
    If Not IsArray(expectedArray) Or Not IsArray(actualArray) Then
        Fail WithMessage("AssertArrayEqual requires two arrays, got " & _
                         Describe(expectedArray) & " and " & Describe(actualArray), message)
    End If
    If LBound(expectedArray) <> LBound(actualArray) Then
        Fail WithMessage("AssertArrayEqual expected LBound " & LBound(expectedArray) & _
                         " but was " & LBound(actualArray), message)
    End If
    If UBound(expectedArray) <> UBound(actualArray) Then
        Fail WithMessage("AssertArrayEqual expected UBound " & UBound(expectedArray) & _
                         " but was " & UBound(actualArray), message)
    End If
    For idx = LBound(expectedArray) To UBound(expectedArray)
        If Not ValuesEqual(expectedArray(idx), actualArray(idx)) Then
            Fail WithMessage("AssertArrayEqual at index " & idx & " expected " & _
                             Describe(expectedArray(idx)) & " but was " & Describe(actualArray(idx)), message)
        End If
    Next idx
End Sub

' Call right after the statement expected to raise, while On Error Resume Next is active.
Public Sub AssertRaised(ByVal expectedNumber As Long, Optional ByVal message As String)
    Dim actualNumber As Long
    Dim actualDesc As String
    actualNumber = Err.Number
    actualDesc = Err.Description
    Err.Clear
    If actualNumber <> expectedNumber Then
        Fail WithMessage("AssertRaised expected error " & expectedNumber & " but was " & _
                         actualNumber & " (" & actualDesc & ")", message)
    End If
End Sub

' ===== File helpers =====

' Creates a fresh folder under the temp path. RunAllTests deletes it after the test.
Public Function NewTempDir() As String
    Dim path As String
    path = Net.Path.GetTempPath() & "dotnetvba-" & NewFso().GetTempName()
    NewFso().CreateFolder path
    mTempDir = path
    NewTempDir = path
End Function

Public Sub WriteBytes(ByVal path As String, ByRef bytes() As Byte)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    If UBound(bytes) >= LBound(bytes) Then stm.Write bytes
    stm.SaveToFile path, 2             ' adSaveCreateOverWrite
    stm.Close
End Sub

Public Function ReadBytes(ByVal path As String) As Byte()
    Dim stm As Object
    Dim data As Variant
    Dim empty0() As Byte
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    data = stm.Read(-1)                ' adReadAll
    stm.Close
    If IsNull(data) Then
        ReDim empty0(0 To -1)          ' empty file: Read returns Null
        ReadBytes = empty0
    Else
        ReadBytes = data
    End If
End Function

' Fixture writer: text mode, BOM is whatever ADODB emits for the charset.
Public Sub WriteTextRaw(ByVal path As String, ByVal text As String, ByVal charset As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                       ' adTypeText
    stm.Charset = charset
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2             ' adSaveCreateOverWrite
    stm.Close
End Sub

' ===== Tests =====

Public Sub Test_Facade_Version()
    AssertEqual "1.0.0", Net.Version
    AssertTrue Net Is Net, "Net() must return the same instance"
End Sub

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

Public Sub Test_Path_EdgeCases()
    AssertEqual "", Net.Path.GetDirectoryName("C:\")
    AssertEqual "", Net.Path.GetDirectoryName("a.txt")
    AssertEqual "\b", Net.Path.Combine("C:\a", "\b")
    AssertEqual "C:\a\b", Net.Path.Combine("C:\a\", "b")
    AssertEqual "C:\a\b\", Net.Path.Combine("C:\a", "b\")
    AssertEqual ".gitignore", Net.Path.GetExtension(".gitignore")
    AssertEqual "", Net.Path.GetFileNameWithoutExtension(".gitignore")
    AssertEqual "\\server\share\a", Net.Path.GetDirectoryName("\\server\share\a\b.txt")
    AssertEqual "c.txt", Net.Path.GetFileName("C:/a/b/c.txt")
    AssertEqual "C:/", Net.Path.GetPathRoot("C:/a/b")
    AssertEqual "//server/share", Net.Path.GetPathRoot("//server/share/a")
    AssertEqual "//server/share/a", Net.Path.GetDirectoryName("//server/share/a/b.txt")
    AssertEqual NewFso().GetAbsolutePathName(CurDir$) & "\y", Net.Path.GetFullPath("x\..\y")
End Sub

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

Public Sub Test_Directory_SetCurrentDirectoryRelative()
    Dim dir As String: dir = NewTempDir()
    Net.Directory.CreateDirectory dir & "\sub"
    Dim saved As String: saved = Net.Directory.GetCurrentDirectory()
    Net.Directory.SetCurrentDirectory dir
    Net.Directory.SetCurrentDirectory "sub"
    Dim after As String: after = Net.Directory.GetCurrentDirectory()
    Net.Directory.SetCurrentDirectory saved
    AssertEqual dir & "\sub", after
End Sub

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
