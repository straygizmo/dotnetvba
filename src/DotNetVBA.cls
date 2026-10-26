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

Public Property Get Version() As String
    Version = VERSION_STRING
End Property

Public Property Get Path() As DotNetPath
    If mPath Is Nothing Then Set mPath = New DotNetPath
    Set Path = mPath
End Property
