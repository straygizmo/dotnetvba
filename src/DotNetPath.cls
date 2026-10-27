VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.IO.Path. Obtain it with Net.Path.

Public Function Combine(ParamArray parts() As Variant) As String
    Dim i As Long
    Dim part As String
    Dim result As String
    For i = LBound(parts) To UBound(parts)
        part = CStr(parts(i))
        If Len(part) > 0 Then
            If Len(result) = 0 Or IsPathRooted(part) Then
                result = part
            ElseIf IsSeparator(Right$(result, 1)) Or IsSeparator(Left$(part, 1)) Then
                result = result & part
            Else
                result = result & "\" & part
            End If
        End If
    Next i
    Combine = result
End Function

Public Function GetFileName(ByVal path As String) As String
    GetFileName = Mid$(path, LastSeparator(path) + 1)
End Function

Public Function GetFileNameWithoutExtension(ByVal path As String) As String
    Dim name As String
    name = GetFileName(path)
    GetFileNameWithoutExtension = Left$(name, Len(name) - Len(GetExtension(name)))
End Function

Public Function GetExtension(ByVal path As String) As String
    Dim name As String
    Dim dotPos As Long
    name = GetFileName(path)
    dotPos = InStrRev(name, ".")
    If dotPos = 0 Or dotPos = Len(name) Then Exit Function
    GetExtension = Mid$(name, dotPos)
End Function

Public Function GetDirectoryName(ByVal path As String) As String
    Dim sepPos As Long
    Dim root As String
    Dim dirName As String
    sepPos = LastSeparator(path)
    If sepPos = 0 Then Exit Function
    root = GetPathRoot(path)
    If Len(path) <= Len(root) Then Exit Function
    dirName = Left$(path, sepPos - 1)
    If Len(dirName) < Len(root) Then dirName = root
    GetDirectoryName = dirName
End Function

Public Function ChangeExtension(ByVal path As String, ByVal extension As String) As String
    Dim baseName As String
    baseName = Left$(path, Len(path) - Len(GetExtension(path)))
    If Len(extension) = 0 Then
        ChangeExtension = baseName
    ElseIf Left$(extension, 1) = "." Then
        ChangeExtension = baseName & extension
    Else
        ChangeExtension = baseName & "." & extension
    End If
End Function

Public Function HasExtension(ByVal path As String) As Boolean
    HasExtension = Len(GetExtension(path)) > 0
End Function

Public Function IsPathRooted(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    IsPathRooted = IsSeparator(Left$(path, 1)) Or Mid$(path, 2, 1) = ":"
End Function

Public Function GetPathRoot(ByVal path As String) As String
    Dim sepPos As Long
    If Len(path) = 0 Then Exit Function
    If IsSeparator(Left$(path, 1)) And IsSeparator(Mid$(path, 2, 1)) Then
        sepPos = NextSeparator(path, 3)
        If sepPos = 0 Then
            GetPathRoot = path
            Exit Function
        End If
        sepPos = NextSeparator(path, sepPos + 1)
        If sepPos = 0 Then
            GetPathRoot = path
        Else
            GetPathRoot = Left$(path, sepPos - 1)
        End If
    ElseIf Mid$(path, 2, 1) = ":" Then
        If IsSeparator(Mid$(path, 3, 1)) Then
            GetPathRoot = Left$(path, 3)
        Else
            GetPathRoot = Left$(path, 2)
        End If
    ElseIf IsSeparator(Left$(path, 1)) Then
        GetPathRoot = Left$(path, 1)
    End If
End Function

Public Function GetFullPath(ByVal path As String) As String
    GetFullPath = NewFso().GetAbsolutePathName(path)
End Function

Public Function GetTempPath() As String
    GetTempPath = NewFso().GetSpecialFolder(2).Path & "\"
End Function

' Creates a new zero-byte file with a unique name; never touches an existing file.
Public Function GetTempFileName() As String
    Dim fso As Object
    Dim fullName As String
    Dim fileNum As Long
    Set fso = NewFso()
    Do
        fullName = GetTempPath() & fso.GetTempName()
    Loop While fso.FileExists(fullName)
    On Error GoTo Failed
    fileNum = FreeFile
    Open fullName For Output As #fileNum
    Close #fileNum
    GetTempFileName = fullName
    Exit Function
Failed:
    ThrowDotNet DotNetIO, "DotNetPath.GetTempFileName", "Unable to create a temporary file: " & Err.Description
End Function

Public Function GetRandomFileName() As String
    GetRandomFileName = NewFso().GetTempName()
End Function

Public Property Get DirectorySeparatorChar() As String
    DirectorySeparatorChar = "\"
End Property

Private Function IsSeparator(ByVal ch As String) As Boolean
    IsSeparator = (ch = "\" Or ch = "/")
End Function

Private Function LastSeparator(ByVal path As String) As Long
    Dim a As Long
    Dim b As Long
    a = InStrRev(path, "\")
    b = InStrRev(path, "/")
    If a > b Then LastSeparator = a Else LastSeparator = b
End Function

' First separator at or after startPos, or 0.
Private Function NextSeparator(ByVal path As String, ByVal startPos As Long) As Long
    Dim a As Long
    Dim b As Long
    a = InStr(startPos, path, "\")
    b = InStr(startPos, path, "/")
    If a = 0 Then
        NextSeparator = b
    ElseIf b = 0 Then
        NextSeparator = a
    ElseIf a < b Then
        NextSeparator = a
    Else
        NextSeparator = b
    End If
End Function
