Attribute VB_Name = "TestAddInConsumer"
Option Explicit

' Add-in boundary test. Imported into a workbook that only holds a reference
' to the built DotNetVBA.xlam, so it uses nothing but the add-in's public API
' and carries its own assert (the test harness is not available here).

Private mLog As Object
Private mFailCount As Long
Private mTotalCount As Long

' Runs the boundary checks, writes PASS/FAIL lines plus a TOTAL line to
' outputPath (UTF-16) and returns the failure count.
Public Function RunAddInConsumerTest(ByVal outputPath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mLog = fso.CreateTextFile(outputPath, True, True)
    mFailCount = 0
    mTotalCount = 0

    Record "AddIn_Version", CheckVersion()
    Record "AddIn_FileAndDirectoryExists", CheckExists()
    Record "AddIn_ListDefaultMember", CheckListDefaultMember()
    Record "AddIn_ListForEach", CheckListForEach()
    Record "AddIn_SpecialFolderEnum", CheckSpecialFolderEnum()
    Record "AddIn_ErrorNumber", CheckErrorNumber()

    mLog.WriteLine "TOTAL " & mTotalCount & " FAILED " & mFailCount
    mLog.Close
    Set mLog = Nothing
    RunAddInConsumerTest = mFailCount
End Function

' Assert: returns failMessage when the condition does not hold, else "".
Private Function Expect(ByVal condition As Boolean, ByVal failMessage As String) As String
    If Not condition Then Expect = failMessage
End Function

Private Sub Record(ByVal testName As String, ByVal errText As String)
    mTotalCount = mTotalCount + 1
    If errText = vbNullString Then
        mLog.WriteLine "PASS " & testName
    Else
        mFailCount = mFailCount + 1
        mLog.WriteLine "FAIL " & testName & ": " & errText
    End If
End Sub

Private Function Unexpected() As String
    Unexpected = "unexpected error " & Err.Number & ": " & Err.Description
End Function

Private Function CheckVersion() As String
    On Error GoTo Failed
    CheckVersion = Expect(Net.Version = "1.0.0", "Net.Version = '" & Net.Version & "'")
    Exit Function
Failed:
    CheckVersion = Unexpected()
End Function

Private Function CheckExists() As String
    Dim tempPath As String
    On Error GoTo Failed
    tempPath = Net.Path.GetTempPath()
    CheckExists = Expect(Net.File.Exists(tempPath) = False, "File.Exists(temp dir) is True")
    If CheckExists = vbNullString Then
        CheckExists = Expect(Net.Directory.Exists(tempPath) = True, "Directory.Exists(temp dir) is False")
    End If
    Exit Function
Failed:
    CheckExists = Unexpected()
End Function

Private Function CheckListDefaultMember() As String
    Dim l As DotNetList
    On Error GoTo Failed
    Set l = Net.NewList()
    l.Add "a"
    l.Add "b"
    CheckListDefaultMember = Expect(l(1) = "b", "l(1) = '" & l(1) & "'")
    Exit Function
Failed:
    CheckListDefaultMember = Unexpected()
End Function

Private Function CheckListForEach() As String
    Dim l As DotNetList
    Dim v As Variant
    Dim joined As String
    On Error GoTo Failed
    Set l = Net.NewList()
    l.Add "a"
    l.Add "b"
    For Each v In l
        joined = joined & v
    Next v
    CheckListForEach = Expect(joined = "ab", "For Each gave '" & joined & "'")
    Exit Function
Failed:
    CheckListForEach = Unexpected()
End Function

Private Function CheckSpecialFolderEnum() As String
    Dim folderPath As String
    On Error GoTo Failed
    folderPath = Net.Environment.GetFolderPath(SpecialFolder_Desktop)
    CheckSpecialFolderEnum = Expect(Len(folderPath) > 0, "GetFolderPath(SpecialFolder_Desktop) is empty")
    Exit Function
Failed:
    CheckSpecialFolderEnum = Unexpected()
End Function

Private Function CheckErrorNumber() As String
    Dim errNum As Long
    Dim errDesc As String
    On Error Resume Next
    Net.File.ReadAllText ""
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo 0
    CheckErrorNumber = Expect(errNum = vbObjectError + 1001, _
        "Err.Number = " & errNum & " (" & errDesc & "), expected vbObjectError + 1001")
End Function
