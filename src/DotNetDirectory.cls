VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.IO.Directory. Obtain it with Net.Directory.

' Wide-character Win32 calls: CurDir$/ChDir go through the ANSI code page.
Private Declare PtrSafe Function GetCurrentDirectoryW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As LongPtr) As Long
Private Declare PtrSafe Function SetCurrentDirectoryW Lib "kernel32" (ByVal lpPathName As LongPtr) As Long

Private mPathHelper As DotNetPath

Public Function Exists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    On Error GoTo Fail
    Exists = NewFso().FolderExists(TrimPath(path))
    Exit Function
Fail:
    Exists = False
End Function

Public Sub CreateDirectory(ByVal path As String)
    On Error GoTo Fail
    CreateRecursive NewFso(), TrimPath(RequirePath(path, "CreateDirectory"))
    Exit Sub
Fail:
    Rethrow "DotNetDirectory.CreateDirectory"
End Sub

Public Sub Delete(ByVal path As String, Optional ByVal recursive As Boolean = False)
    Dim fso As Object
    Dim target As String
    target = TrimPath(RequirePath(path, "Delete"))
    Set fso = NewFso()
    On Error GoTo Fail
    If Not fso.FolderExists(target) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetDirectory.Delete", "Could not find a part of the path '" & path & "'."
    End If
    If Not recursive Then
        If fso.GetFolder(target).Files.Count > 0 Or fso.GetFolder(target).SubFolders.Count > 0 Then
            ThrowDotNet DotNetIO, "DotNetDirectory.Delete", "The directory is not empty: '" & path & "'."
        End If
    End If
    fso.DeleteFolder target, True
    Exit Sub
Fail:
    Rethrow "DotNetDirectory.Delete"
End Sub

Public Sub Move(ByVal sourceDirName As String, ByVal destDirName As String)
    Dim fso As Object
    Dim src As String
    Dim dest As String
    src = TrimPath(RequirePath(sourceDirName, "Move"))
    dest = TrimPath(RequirePath(destDirName, "Move"))
    Set fso = NewFso()
    On Error GoTo Fail
    If Not fso.FolderExists(src) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetDirectory.Move", "Could not find a part of the path '" & sourceDirName & "'."
    End If
    If fso.FolderExists(dest) Or fso.FileExists(dest) Then
        ThrowDotNet DotNetIO, "DotNetDirectory.Move", "Cannot create '" & destDirName & "' because a file or directory with the same name already exists."
    End If
    fso.MoveFolder src, dest
    Exit Sub
Fail:
    Rethrow "DotNetDirectory.Move"
End Sub

Public Function GetFiles(ByVal path As String, Optional ByVal searchPattern As String = "*", Optional ByVal recursive As Boolean = False) As String()
    GetFiles = Enumerate(path, searchPattern, recursive, True, "GetFiles")
End Function

Public Function GetDirectories(ByVal path As String, Optional ByVal searchPattern As String = "*", Optional ByVal recursive As Boolean = False) As String()
    GetDirectories = Enumerate(path, searchPattern, recursive, False, "GetDirectories")
End Function

Public Function GetParent(ByVal path As String) As String
    Dim target As String
    target = TrimPath(RequirePath(path, "GetParent"))
    GetParent = PathHelper().GetDirectoryName(target)
End Function

Public Function GetCurrentDirectory() As String
    Dim buf As String
    Dim size As Long
    Dim ret As Long
    size = 32767
    Do
        buf = String$(size, vbNullChar)
        ret = GetCurrentDirectoryW(size, StrPtr(buf))
        If ret = 0 Then
            ThrowDotNet DotNetIO, "DotNetDirectory.GetCurrentDirectory", "GetCurrentDirectoryW failed (error " & CStr(Err.LastDllError) & ")."
        End If
        If ret < size Then Exit Do
        size = ret + 1
    Loop
    GetCurrentDirectory = Left$(buf, ret)
End Function

Public Sub SetCurrentDirectory(ByVal path As String)
    Dim target As String
    target = RequirePath(path, "SetCurrentDirectory")
    On Error GoTo Fail
    target = TrimPath(NewFso().GetAbsolutePathName(target))
    If Not NewFso().FolderExists(target) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetDirectory.SetCurrentDirectory", "Could not find a part of the path '" & path & "'."
    End If
    If SetCurrentDirectoryW(StrPtr(target)) = 0 Then
        ThrowDotNet DotNetIO, "DotNetDirectory.SetCurrentDirectory", "Could not set the current directory to '" & path & "' (error " & CStr(Err.LastDllError) & ")."
    End If
    Exit Sub
Fail:
    Rethrow "DotNetDirectory.SetCurrentDirectory"
End Sub

Public Function GetCreationTime(ByVal path As String) As Date
    GetCreationTime = FolderOf(path, "GetCreationTime").DateCreated
End Function

Public Function GetLastWriteTime(ByVal path As String) As Date
    GetLastWriteTime = FolderOf(path, "GetLastWriteTime").DateLastModified
End Function

' ----- helpers -----

Private Function PathHelper() As DotNetPath
    If mPathHelper Is Nothing Then Set mPathHelper = New DotNetPath
    Set PathHelper = mPathHelper
End Function

Private Function RequirePath(ByVal path As String, ByVal memberName As String) As String
    If Len(path) = 0 Then
        ThrowDotNet DotNetArgument, "DotNetDirectory." & memberName, "The path is empty. (Parameter 'path')"
    End If
    RequirePath = path
End Function

' Strips one trailing separator unless the path is a root.
Private Function TrimPath(ByVal path As String) As String
    Dim lastChar As String
    lastChar = Right$(path, 1)
    If (lastChar = "\" Or lastChar = "/") And Len(path) > Len(PathHelper().GetPathRoot(path)) Then
        path = Left$(path, Len(path) - 1)
    End If
    TrimPath = path
End Function

Private Sub CreateRecursive(ByVal fso As Object, ByVal target As String)
    Dim parent As String
    If fso.FolderExists(target) Then Exit Sub
    If fso.FileExists(target) Then
        ThrowDotNet DotNetIO, "DotNetDirectory.CreateDirectory", "Cannot create '" & target & "' because a file with the same name already exists."
    End If
    parent = fso.GetParentFolderName(target)
    If Len(parent) > 0 Then CreateRecursive fso, parent
    fso.CreateFolder target
End Sub

Private Function FolderOf(ByVal path As String, ByVal memberName As String) As Object
    Dim fso As Object
    Dim target As String
    target = TrimPath(RequirePath(path, memberName))
    Set fso = NewFso()
    On Error GoTo Fail
    If Not fso.FolderExists(target) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetDirectory." & memberName, "Could not find a part of the path '" & path & "'."
    End If
    Set FolderOf = fso.GetFolder(target)
    Exit Function
Fail:
    Rethrow "DotNetDirectory." & memberName
End Function

Private Function Enumerate(ByVal path As String, ByVal searchPattern As String, ByVal recursive As Boolean, ByVal wantFiles As Boolean, ByVal memberName As String) As String()
    Dim fso As Object
    Dim target As String
    Dim found As Collection
    Dim result() As String
    Dim i As Long
    target = TrimPath(RequirePath(path, memberName))
    Set fso = NewFso()
    Set found = New Collection
    On Error GoTo Fail
    If Not fso.FolderExists(target) Then
        ThrowDotNet DotNetDirectoryNotFound, "DotNetDirectory." & memberName, "Could not find a part of the path '" & path & "'."
    End If
    Collect fso, target, searchPattern, recursive, wantFiles, found
    If found.Count = 0 Then
        Enumerate = EmptyStringArray()
        Exit Function
    End If
    ReDim result(0 To found.Count - 1)
    For i = 1 To found.Count
        result(i - 1) = found(i)
    Next i
    Enumerate = result
    Exit Function
Fail:
    Rethrow "DotNetDirectory." & memberName
End Function

Private Sub Collect(ByVal fso As Object, ByVal folderPath As String, ByVal searchPattern As String, ByVal recursive As Boolean, ByVal wantFiles As Boolean, ByVal found As Collection)
    Dim folder As Object
    Dim item As Object
    Dim basePath As String
    basePath = folderPath
    If Right$(basePath, 1) = "\" Or Right$(basePath, 1) = "/" Then basePath = Left$(basePath, Len(basePath) - 1)
    Set folder = fso.GetFolder(folderPath)
    If wantFiles Then
        For Each item In folder.Files
            If MatchesPattern(item.Name, searchPattern) Then found.Add basePath & "\" & item.Name
        Next item
    Else
        For Each item In folder.SubFolders
            If MatchesPattern(item.Name, searchPattern) Then found.Add basePath & "\" & item.Name
        Next item
    End If
    If recursive Then
        For Each item In folder.SubFolders
            Collect fso, basePath & "\" & item.Name, searchPattern, True, wantFiles, found
        Next item
    End If
End Sub

' Re-raises DotNetVBA errors unchanged; wraps anything else (e.g. raw FSO errors) as DotNetIO.
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
