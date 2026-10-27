VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' System.Threading.Thread. Obtain it with Net.Thread.

Private Declare PtrSafe Sub SleepApi Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

' Suspends the calling thread. 0 yields the time slice.
Public Sub Sleep(ByVal milliseconds As Long)
    If milliseconds < 0 Then
        ThrowDotNet DotNetArgumentOutOfRange, "DotNetThread.Sleep", "Number must be either non-negative or -1. (Parameter 'milliseconds')"
    End If
    SleepApi milliseconds
End Sub
