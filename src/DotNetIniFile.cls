VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' INI file access over the Win32 private-profile API. Obtain it with Net.IniFile.
' Note: the W API writes ANSI unless the file already exists as UTF-16 with BOM.

Private Declare PtrSafe Function GetPrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpDefault As LongPtr, ByVal lpReturnedString As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function WritePrivateProfileStringW Lib "kernel32" (ByVal lpAppName As LongPtr, ByVal lpKeyName As LongPtr, ByVal lpString As LongPtr, ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function GetPrivateProfileSectionNamesW Lib "kernel32" (ByVal lpszReturnBuffer As LongPtr, ByVal nSize As Long, ByVal lpFileName As LongPtr) As Long

Private Const INITIAL_BUFFER As Long = 32767

' Returns defaultValue when the file, section or key is missing.
Public Function ReadString(ByVal path As String, ByVal section As String, ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    Dim full As String
    Dim sentinel As String
    Dim buf As String
    Dim n As Long
    Dim size As Long
    Dim result As String
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.ReadString")
    RequireName section, "section", "DotNetIniFile.ReadString"
    RequireName keyName, "keyName", "DotNetIniFile.ReadString"
    sentinel = ChrW$(1) & "DNV" & ChrW$(2)
    size = INITIAL_BUFFER
    Do
        buf = String$(size, vbNullChar)
        n = GetPrivateProfileStringW(StrPtr(section), StrPtr(keyName), StrPtr(sentinel), StrPtr(buf), size, StrPtr(full))
        If n < size - 1 Then Exit Do
        size = size * 2
    Loop
    result = Left$(buf, n)
    If result = sentinel Then result = defaultValue
    ReadString = result
    Exit Function
Fail:
    Rethrow "DotNetIniFile.ReadString"
End Function

Public Sub WriteString(ByVal path As String, ByVal section As String, ByVal keyName As String, ByVal value As String)
    Dim full As String
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.WriteString")
    RequireName section, "section", "DotNetIniFile.WriteString"
    RequireName keyName, "keyName", "DotNetIniFile.WriteString"
    RequireParent full, "DotNetIniFile.WriteString"
    CheckWrite WritePrivateProfileStringW(StrPtr(section), StrPtr(keyName), StrPtr(value), StrPtr(full))
    Exit Sub
Fail:
    Rethrow "DotNetIniFile.WriteString"
End Sub

Public Sub DeleteKey(ByVal path As String, ByVal section As String, ByVal keyName As String)
    Dim full As String
    Dim nullPtr As LongPtr
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.DeleteKey")
    RequireName section, "section", "DotNetIniFile.DeleteKey"
    RequireName keyName, "keyName", "DotNetIniFile.DeleteKey"
    RequireParent full, "DotNetIniFile.DeleteKey"
    CheckWrite WritePrivateProfileStringW(StrPtr(section), StrPtr(keyName), nullPtr, StrPtr(full))
    Exit Sub
Fail:
    Rethrow "DotNetIniFile.DeleteKey"
End Sub

Public Sub DeleteSection(ByVal path As String, ByVal section As String)
    Dim full As String
    Dim nullPtr As LongPtr
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.DeleteSection")
    RequireName section, "section", "DotNetIniFile.DeleteSection"
    RequireParent full, "DotNetIniFile.DeleteSection"
    CheckWrite WritePrivateProfileStringW(StrPtr(section), nullPtr, nullPtr, StrPtr(full))
    Exit Sub
Fail:
    Rethrow "DotNetIniFile.DeleteSection"
End Sub

Public Function GetSections(ByVal path As String) As String()
    Dim full As String
    Dim buf As String
    Dim n As Long
    Dim size As Long
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.GetSections")
    size = INITIAL_BUFFER
    Do
        buf = String$(size, vbNullChar)
        n = GetPrivateProfileSectionNamesW(StrPtr(buf), size, StrPtr(full))
        If n < size - 2 Then Exit Do
        size = size * 2
    Loop
    GetSections = SplitList(buf, n)
    Exit Function
Fail:
    Rethrow "DotNetIniFile.GetSections"
End Function

Public Function GetKeys(ByVal path As String, ByVal section As String) As String()
    Dim full As String
    Dim buf As String
    Dim n As Long
    Dim size As Long
    Dim nullPtr As LongPtr
    On Error GoTo Fail
    full = ResolvePath(path, "DotNetIniFile.GetKeys")
    RequireName section, "section", "DotNetIniFile.GetKeys"
    size = INITIAL_BUFFER
    Do
        buf = String$(size, vbNullChar)
        n = GetPrivateProfileStringW(StrPtr(section), nullPtr, nullPtr, StrPtr(buf), size, StrPtr(full))
        If n < size - 2 Then Exit Do
        size = size * 2
    Loop
    GetKeys = SplitList(buf, n)
    Exit Function
Fail:
    Rethrow "DotNetIniFile.GetKeys"
End Function

' ----- helpers -----

' Returns the absolute path; the API would treat a bare name as relative to the Windows folder.
Private Function ResolvePath(ByVal path As String, ByVal source As String) As String
    If Len(path) = 0 Then
        ThrowDotNet DotNetArgument, source, "The path is empty. (Parameter 'path')"
    End If
    ResolvePath = NewFso().GetAbsolutePathName(path)
End Function

Private Sub RequireName(ByVal value As String, ByVal paramName As String, ByVal source As String)
    If Len(value) = 0 Then
        ThrowDotNet DotNetArgument, source, "The value is empty. (Parameter '" & paramName & "')"
    End If
End Sub

Private Sub RequireParent(ByVal full As String, ByVal source As String)
    Dim fso As Object
    Dim parent As String
    Set fso = NewFso()
    parent = fso.GetParentFolderName(full)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then
            ThrowDotNet DotNetDirectoryNotFound, source, "Could not find a part of the path '" & full & "'."
        End If
    End If
End Sub

Private Sub CheckWrite(ByVal apiResult As Long)
    Dim lastErr As Long
    lastErr = Err.LastDllError
    If apiResult = 0 Then
        ThrowDotNet DotNetIO, "DotNetIniFile", "WritePrivateProfileString failed (error " & CStr(lastErr) & ")."
    End If
End Sub

' Splits "a\0b\0\0" (first n characters of buf) into a String array.
Private Function SplitList(ByVal buf As String, ByVal n As Long) As String()
    Dim text As String
    text = Left$(buf, n)
    Do While Right$(text, 1) = vbNullChar
        text = Left$(text, Len(text) - 1)
    Loop
    If Len(text) = 0 Then
        SplitList = EmptyStringArray()
    Else
        SplitList = Split(text, vbNullChar)
    End If
End Function

' Re-raises DotNetVBA errors unchanged; wraps anything else as DotNetIO.
Private Sub Rethrow(ByVal source As String)
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    If errNum >= vbObjectError + DotNetArgument And errNum <= vbObjectError + DotNetNotSupported Then
        Err.Raise errNum, errSrc, errDesc
    End If
    ThrowDotNet DotNetIO, source, errDesc
End Sub
