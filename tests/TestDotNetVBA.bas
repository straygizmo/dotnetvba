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
    Set logFile = fso.CreateTextFile(outputPath, True)
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

' ===== Tests =====

Public Sub Test_Facade_Version()
    AssertEqual "1.0.0", Net.Version
    AssertTrue Net Is Net, "Net() must return the same instance"
End Sub
