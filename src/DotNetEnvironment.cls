VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetEnvironment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.Environment. Obtain it with Net.Environment.

' Wide-character Win32 call: Environ$ goes through the ANSI code page.
Private Declare PtrSafe Function GetEnvironmentVariableW Lib "kernel32" (ByVal lpName As LongPtr, ByVal lpBuffer As LongPtr, ByVal nSize As Long) As Long

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

Public Function GetFolderPath(ByVal folder As SpecialFolder) As String
    Dim result As String
    On Error GoTo Fail
    Select Case folder
        Case SpecialFolder_Desktop
            result = ShellFolder("Desktop")
        Case SpecialFolder_MyDocuments
            result = ShellFolder("MyDocuments")
        Case SpecialFolder_ApplicationData
            result = ExpandEnvironmentVariables("%APPDATA%")
        Case SpecialFolder_LocalApplicationData
            result = ExpandEnvironmentVariables("%LOCALAPPDATA%")
        Case SpecialFolder_UserProfile
            result = ExpandEnvironmentVariables("%USERPROFILE%")
        Case SpecialFolder_ProgramFiles
            result = ExpandEnvironmentVariables("%ProgramFiles%")
        Case SpecialFolder_Startup
            result = ShellFolder("Startup")
        Case SpecialFolder_StartMenu
            result = ShellFolder("StartMenu")
        Case SpecialFolder_Templates
            result = ShellFolder("Templates")
        Case SpecialFolder_Favorites
            result = ShellFolder("Favorites")
        Case SpecialFolder_Recent
            result = ShellFolder("Recent")
        Case SpecialFolder_SendTo
            result = ShellFolder("SendTo")
        Case Else
            ThrowDotNet DotNetArgument, "DotNetEnvironment.GetFolderPath", "Value " & CStr(folder) & " is not a valid SpecialFolder. (Parameter 'folder')"
    End Select
    If Right$(result, 1) = "\" And Len(result) > 3 Then result = Left$(result, Len(result) - 1)
    GetFolderPath = result
    Exit Function
Fail:
    Rethrow "DotNetEnvironment.GetFolderPath"
End Function

Public Property Get NewLine() As String
    NewLine = vbCrLf
End Property

Public Property Get UserName() As String
    UserName = GetEnvironmentVariable("USERNAME")
End Property

Public Property Get MachineName() As String
    MachineName = GetEnvironmentVariable("COMPUTERNAME")
End Property

Public Property Get CurrentDirectory() As String
    CurrentDirectory = Net.Directory.GetCurrentDirectory()
End Property

Public Property Let CurrentDirectory(ByVal value As String)
    Net.Directory.SetCurrentDirectory value
End Property

' Returns "" when the variable is not set.
Public Function GetEnvironmentVariable(ByVal variableName As String) As String
    If Len(variableName) = 0 Then
        ThrowDotNet DotNetArgument, "DotNetEnvironment.GetEnvironmentVariable", "The variable name is empty. (Parameter 'variable')"
    End If
    GetEnvironmentVariable = ReadVariable(variableName)
End Function

' Unknown %NAME% tokens are left untouched.
Public Function ExpandEnvironmentVariables(ByVal text As String) As String
    On Error GoTo Fail
    ExpandEnvironmentVariables = CreateObject("WScript.Shell").ExpandEnvironmentStrings(text)
    Exit Function
Fail:
    Rethrow "DotNetEnvironment.ExpandEnvironmentVariables"
End Function

' DotNetVBA is a 64-bit only library.
Public Property Get Is64BitProcess() As Boolean
    Is64BitProcess = True
End Property

' ----- helpers -----

' Returns "" when the variable is not set (or is set to an empty value).
Private Function ReadVariable(ByVal variableName As String) As String
    Dim buf As String
    Dim size As Long
    Dim ret As Long
    size = 1024
    Do
        buf = String$(size, vbNullChar)
        ret = GetEnvironmentVariableW(StrPtr(variableName), StrPtr(buf), size)
        If ret = 0 Then Exit Function
        If ret < size Then Exit Do
        size = ret
    Loop
    ReadVariable = Left$(buf, ret)
End Function

Private Function ShellFolder(ByVal folderName As String) As String
    ShellFolder = CStr(CreateObject("WScript.Shell").SpecialFolders(folderName))
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
