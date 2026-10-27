VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Facade over the DotNetVBA classes. Obtain it with Net().

Private Const VERSION_STRING As String = "1.0.0"
Private mPath As DotNetPath
Private mDirectory As DotNetDirectory
Private mFile As DotNetFile
Private mEnvironment As DotNetEnvironment
Private mIniFile As DotNetIniFile

Public Property Get Version() As String
    Version = VERSION_STRING
End Property

Public Property Get Path() As DotNetPath
    If mPath Is Nothing Then Set mPath = New DotNetPath
    Set Path = mPath
End Property

Public Property Get Directory() As DotNetDirectory
    If mDirectory Is Nothing Then Set mDirectory = New DotNetDirectory
    Set Directory = mDirectory
End Property

Public Property Get File() As DotNetFile
    If mFile Is Nothing Then Set mFile = New DotNetFile
    Set File = mFile
End Property

Public Property Get Environment() As DotNetEnvironment
    If mEnvironment Is Nothing Then Set mEnvironment = New DotNetEnvironment
    Set Environment = mEnvironment
End Property

Public Property Get IniFile() As DotNetIniFile
    If mIniFile Is Nothing Then Set mIniFile = New DotNetIniFile
    Set IniFile = mIniFile
End Property

' Returns a new, empty DotNetList on every call.
Public Function NewList() As DotNetList
    Set NewList = New DotNetList
End Function
