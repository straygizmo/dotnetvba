Attribute VB_Name = "DotNetVBAInternal"
Option Explicit
Option Private Module

' Error numbers raised by DotNetVBA, offset from vbObjectError.
Public Enum DotNetErrorNumber
    DotNetArgument = 1001
    DotNetArgumentOutOfRange = 1002
    DotNetFileNotFound = 1003
    DotNetDirectoryNotFound = 1004
    DotNetIO = 1005
    DotNetInvalidOperation = 1006
    DotNetNotSupported = 1007
End Enum

' Raises vbObjectError + errNumber with a .NET-style exception description.
Public Sub ThrowDotNet(ByVal errNumber As DotNetErrorNumber, ByVal source As String, ByVal message As String)
    Err.Raise vbObjectError + errNumber, "DotNetVBA." & source, ExceptionName(errNumber) & ": " & message
End Sub

Public Function NewFso() As Object
    Set NewFso = CreateObject("Scripting.FileSystemObject")
End Function

' Returns a String array with LBound 0 and UBound -1.
Public Function EmptyStringArray() As String()
    EmptyStringArray = Split(vbNullString)
End Function

Public Function ExceptionName(ByVal errNumber As DotNetErrorNumber) As String
    Select Case errNumber
        Case DotNetArgument
            ExceptionName = "ArgumentException"
        Case DotNetArgumentOutOfRange
            ExceptionName = "ArgumentOutOfRangeException"
        Case DotNetFileNotFound
            ExceptionName = "FileNotFoundException"
        Case DotNetDirectoryNotFound
            ExceptionName = "DirectoryNotFoundException"
        Case DotNetIO
            ExceptionName = "IOException"
        Case DotNetInvalidOperation
            ExceptionName = "InvalidOperationException"
        Case DotNetNotSupported
            ExceptionName = "NotSupportedException"
        Case Else
            ExceptionName = "Exception"
    End Select
End Function
