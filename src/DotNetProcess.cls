VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.Diagnostics.Process. Obtain it with Net.Process.

' Starts a program or opens a file/folder without waiting for it to exit.
' windowStyle is the WScript.Shell.Run style: 1 = normal, 3 = maximized,
' 7 = minimized without focus.
Public Sub Start(ByVal fileName As String, Optional ByVal arguments As String = "", Optional ByVal windowStyle As Long = 1)
    Dim cmd As String
    If Len(fileName) = 0 Then
        ThrowDotNet DotNetArgument, "DotNetProcess.Start", "The file name is empty. (Parameter 'fileName')"
    End If
    cmd = """" & fileName & """"
    If Len(arguments) > 0 Then cmd = cmd & " " & arguments
    On Error GoTo Fail
    CreateObject("WScript.Shell").Run cmd, windowStyle, False
    Exit Sub
Fail:
    ThrowDotNet DotNetIO, "DotNetProcess.Start", Err.Description
End Sub
