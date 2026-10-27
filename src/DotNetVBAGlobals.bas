Attribute VB_Name = "DotNetVBAGlobals"
Option Explicit

Private mInstance As DotNet

' Returns the shared DotNet facade, creating it on first call.
Public Function Net() As DotNet
    If mInstance Is Nothing Then Set mInstance = New DotNet
    Set Net = mInstance
End Function
