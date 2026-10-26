Attribute VB_Name = "DotNetVBAGlobals"
Option Explicit

Private mInstance As DotNetVBA

' Returns the shared DotNetVBA facade, creating it on first call.
Public Function Net() As DotNetVBA
    If mInstance Is Nothing Then Set mInstance = New DotNetVBA
    Set Net = mInstance
End Function
