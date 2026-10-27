VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.IO.File. Obtain it with Net.File.
' Text goes through ADODB.Stream; "utf-8" (the default) is written without a BOM.

Private Const AD_TYPE_BINARY As Long = 1            ' adTypeBinary
Private Const AD_TYPE_TEXT As Long = 2              ' adTypeText
Private Const AD_READ_ALL As Long = -1              ' adReadAll
Private Const AD_SAVE_CREATE_OVERWRITE As Long = 2  ' adSaveCreateOverWrite
Private Const AD_STATE_CLOSED As Long = 0           ' adStateClosed

Public Function Exists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    On Error GoTo Fail
    Exists = NewFso().FileExists(path)
    Exit Function
Fail:
    Exists = False
End Function

Public Sub Copy(ByVal source As String, ByVal dest As String, Optional ByVal overwrite As Boolean = False)
    Dim fso As Object
    RequirePath source, "Copy", "source"
    RequirePath dest, "Copy", "dest"
    Set fso = NewFso()
    On Error GoTo Fail
    RequireFile fso, source, "Copy"
    RequireParent fso, dest, "Copy"
    RequireFreeDest fso, dest, overwrite, "Copy"
    fso.CopyFile source, dest, True
    Exit Sub
Fail:
    Rethrow "DotNetFile.Copy"
End Sub

Public Sub Move(ByVal source As String, ByVal dest As String, Optional ByVal overwrite As Boolean = False)
    Dim fso As Object
    RequirePath source, "Move", "source"
    RequirePath dest, "Move", "dest"
    Set fso = NewFso()
    On Error GoTo Fail
    RequireFile fso, source, "Move"
    RequireParent fso, dest, "Move"
    ' Moving a file onto itself is a no-op (and must not delete it via the overwrite path).
    If StrComp(fso.GetAbsolutePathName(source), fso.GetAbsolutePathName(dest), vbTextCompare) = 0 Then Exit Sub
    RequireFreeDest fso, dest, overwrite, "Move"
    If fso.FileExists(dest) Then fso.DeleteFile dest, False
    fso.MoveFile source, dest
    Exit Sub
Fail:
    Rethrow "DotNetFile.Move"
End Sub

Public Sub Delete(ByVal path As String)
    Dim fso As Object
    RequirePath path, "Delete", "path"
    Set fso = NewFso()
    On Error GoTo Fail
    RequireParent fso, path, "Delete"
    If fso.FileExists(path) Then fso.DeleteFile path, False
    Exit Sub
Fail:
    Rethrow "DotNetFile.Delete"
End Sub

Public Function ReadAllText(ByVal path As String, Optional ByVal encoding As String = "utf-8") As String
    ReadAllText = ReadTextCore(path, encoding, "ReadAllText")
End Function

Public Function ReadAllLines(ByVal path As String, Optional ByVal encoding As String = "utf-8") As String()
    Dim text As String
    text = ReadTextCore(path, encoding, "ReadAllLines")
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    If Right$(text, 1) = vbLf Then text = Left$(text, Len(text) - 1)
    If Len(text) = 0 Then
        ReadAllLines = EmptyStringArray()
    Else
        ReadAllLines = Split(text, vbLf)
    End If
End Function

Public Sub WriteAllText(ByVal path As String, ByVal contents As String, Optional ByVal encoding As String = "utf-8")
    WriteTextCore path, contents, encoding, "WriteAllText"
End Sub

Public Sub WriteAllLines(ByVal path As String, ByVal lines As Variant, Optional ByVal encoding As String = "utf-8")
    RequirePath path, "WriteAllLines", "path"
    WriteTextCore path, JoinLines(lines, "WriteAllLines"), encoding, "WriteAllLines"
End Sub

Public Sub AppendAllText(ByVal path As String, ByVal contents As String, Optional ByVal encoding As String = "utf-8")
    AppendCore path, contents, encoding, "AppendAllText"
End Sub

Public Sub AppendAllLines(ByVal path As String, ByVal lines As Variant, Optional ByVal encoding As String = "utf-8")
    RequirePath path, "AppendAllLines", "path"
    AppendCore path, JoinLines(lines, "AppendAllLines"), encoding, "AppendAllLines"
End Sub

Public Function GetCreationTime(ByVal path As String) As Date
    GetCreationTime = FileOf(path, "GetCreationTime").DateCreated
End Function

Public Function GetLastWriteTime(ByVal path As String) As Date
    GetLastWriteTime = FileOf(path, "GetLastWriteTime").DateLastModified
End Function

Public Function GetLastAccessTime(ByVal path As String) As Date
    GetLastAccessTime = FileOf(path, "GetLastAccessTime").DateLastAccessed
End Function

Public Function GetLength(ByVal path As String) As Double
    GetLength = CDbl(FileOf(path, "GetLength").Size)
End Function

Public Function GetAttributes(ByVal path As String) As Long
    GetAttributes = CLng(FileOf(path, "GetAttributes").Attributes)
End Function

' ----- helpers -----

Private Sub RequirePath(ByVal path As String, ByVal memberName As String, ByVal paramName As String)
    If Len(path) = 0 Then
        ThrowDotNet DotNetArgument, "DotNetFile." & memberName, "The value cannot be an empty string. (Parameter '" & paramName & "')"
    End If
End Sub

Private Sub RequireFile(ByVal fso As Object, ByVal path As String, ByVal memberName As String)
    If Not fso.FileExists(path) Then
        ThrowDotNet DotNetFileNotFound, "DotNetFile." & memberName, "Could not find file '" & path & "'."
    End If
End Sub

Private Sub RequireParent(ByVal fso As Object, ByVal path As String, ByVal memberName As String)
    Dim parent As String
    parent = fso.GetParentFolderName(fso.GetAbsolutePathName(path))
    If Len(parent) = 0 Then Exit Sub
    If Not fso.FolderExists(parent) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetFile." & memberName, "Could not find a part of the path '" & path & "'."
    End If
End Sub

Private Sub RequireFreeDest(ByVal fso As Object, ByVal dest As String, ByVal overwrite As Boolean, ByVal memberName As String)
    If fso.FolderExists(dest) Then
        ThrowDotNet DotNetIO, "DotNetFile." & memberName, "Cannot create '" & dest & "' because a directory with the same name already exists."
    End If
    If fso.FileExists(dest) And Not overwrite Then
        ThrowDotNet DotNetIO, "DotNetFile." & memberName, "The file '" & dest & "' already exists."
    End If
End Sub

Private Function FileOf(ByVal path As String, ByVal memberName As String) As Object
    Dim fso As Object
    RequirePath path, memberName, "path"
    Set fso = NewFso()
    On Error GoTo Fail
    RequireFile fso, path, memberName
    Set FileOf = fso.GetFile(path)
    Exit Function
Fail:
    Rethrow "DotNetFile." & memberName
End Function

' ADODB knows "utf-8" but not the common alias "utf8".
Private Function NormalizeEncoding(ByVal encoding As String) As String
    If LCase$(encoding) = "utf8" Then
        NormalizeEncoding = "utf-8"
    Else
        NormalizeEncoding = encoding
    End If
End Function

' Returns an open text stream using the given charset; an unknown charset raises DotNetArgument.
Private Function OpenTextStream(ByVal encoding As String, ByVal memberName As String) As Object
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = AD_TYPE_TEXT
    On Error GoTo BadCharset
    stream.Charset = encoding
    On Error GoTo 0
    stream.Open
    Set OpenTextStream = stream
    Exit Function
BadCharset:
    ThrowDotNet DotNetArgument, "DotNetFile." & memberName, "'" & encoding & "' is not a supported encoding name. " & Err.Description & " (Parameter 'encoding')"
End Function

Private Function ReadTextCore(ByVal path As String, ByVal encoding As String, ByVal memberName As String) As String
    Dim fso As Object
    Dim stream As Object
    Dim text As Variant
    RequirePath path, memberName, "path"
    Set fso = NewFso()
    On Error GoTo Fail
    RequireFile fso, path, memberName
    Set stream = OpenTextStream(NormalizeEncoding(encoding), memberName)
    stream.LoadFromFile path
    If stream.Size > 0 Then
        text = stream.ReadText(AD_READ_ALL)
        If Not IsNull(text) Then ReadTextCore = text
    End If
    stream.Close
    Exit Function
Fail:
    Rethrow "DotNetFile." & memberName, stream
End Function

' Writes contents in the given charset; for utf-8 the BOM ADODB emits is skipped.
Private Sub WriteTextCore(ByVal path As String, ByVal contents As String, ByVal encoding As String, ByVal memberName As String)
    Dim fso As Object
    Dim textStream As Object
    Dim binaryStream As Object
    Dim head() As Byte
    Dim charsetName As String
    RequirePath path, memberName, "path"
    Set fso = NewFso()
    charsetName = NormalizeEncoding(encoding)
    On Error GoTo Fail
    RequireParent fso, path, memberName
    Set textStream = OpenTextStream(charsetName, memberName)
    textStream.WriteText contents
    textStream.Position = 0
    textStream.Type = AD_TYPE_BINARY
    If LCase$(charsetName) = "utf-8" And textStream.Size >= 3 Then
        head = textStream.Read(3)
        If Not (head(0) = &HEF And head(1) = &HBB And head(2) = &HBF) Then textStream.Position = 0
    End If
    Set binaryStream = CreateObject("ADODB.Stream")
    binaryStream.Type = AD_TYPE_BINARY
    binaryStream.Open
    textStream.CopyTo binaryStream
    binaryStream.SaveToFile path, AD_SAVE_CREATE_OVERWRITE
    binaryStream.Close
    textStream.Close
    Exit Sub
Fail:
    Rethrow "DotNetFile." & memberName, textStream, binaryStream
End Sub

' Rewrites the file as its current text plus contents, keeping one encoding for both halves.
Private Sub AppendCore(ByVal path As String, ByVal contents As String, ByVal encoding As String, ByVal memberName As String)
    RequirePath path, memberName, "path"
    If NewFso().FileExists(path) Then
        contents = ReadTextCore(path, encoding, memberName) & contents
    End If
    WriteTextCore path, contents, encoding, memberName
End Sub

' Joins the elements with CRLF and terminates the last line, as File.WriteAllLines does.
Private Function JoinLines(ByVal lines As Variant, ByVal memberName As String) As String
    Dim parts() As String
    Dim lower As Long
    Dim upper As Long
    Dim i As Long
    If Not IsArray(lines) Then
        ThrowDotNet DotNetArgument, "DotNetFile." & memberName, "The value must be an array. (Parameter 'lines')"
    End If
    On Error GoTo Unallocated
    lower = LBound(lines)
    upper = UBound(lines)
    On Error GoTo 0
    If upper < lower Then Exit Function
    ReDim parts(lower To upper)
    For i = lower To upper
        parts(i) = CStr(lines(i))
    Next i
    JoinLines = Join(parts, vbCrLf) & vbCrLf
    Exit Function
Unallocated:
    JoinLines = vbNullString
End Function

Private Sub CloseQuietly(ByVal stream As Object)
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State <> AD_STATE_CLOSED Then stream.Close
    End If
End Sub

' Re-raises DotNetVBA errors unchanged; wraps anything else (e.g. raw FSO/ADODB errors) as DotNetIO.
' Closes the given streams first so no file handle is left open.
Private Sub Rethrow(ByVal source As String, Optional ByVal stream1 As Object, Optional ByVal stream2 As Object)
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    CloseQuietly stream1
    CloseQuietly stream2
    If errNum >= vbObjectError + DotNetArgument And errNum <= vbObjectError + DotNetNotSupported Then
        Err.Raise errNum, errSrc, errDesc
    End If
    ThrowDotNet DotNetIO, source, errDesc
End Sub
