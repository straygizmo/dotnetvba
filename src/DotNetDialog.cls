VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Office FileDialog wrappers. Obtain it with Net.Dialog.
' Application resolves late-bound to the host (Excel, Word or Access).

Private Const MSO_FILE_DIALOG_SAVE_AS As Long = 2          ' msoFileDialogSaveAs
Private Const MSO_FILE_DIALOG_FILE_PICKER As Long = 3      ' msoFileDialogFilePicker
Private Const MSO_FILE_DIALOG_FOLDER_PICKER As Long = 4    ' msoFileDialogFolderPicker

' Shows an Open dialog; returns the chosen path, or "" when cancelled.
Public Function OpenFile(Optional ByVal title As String = "", Optional ByVal filterName As String = "", Optional ByVal filterPattern As String = "", Optional ByVal initialDirectory As String = "") As String
    Dim picked() As String
    Validate "DotNetDialog.OpenFile", filterName, filterPattern, initialDirectory
    picked = ShowDialog("DotNetDialog.OpenFile", MSO_FILE_DIALOG_FILE_PICKER, title, filterName, filterPattern, initialDirectory, False)
    If UBound(picked) >= LBound(picked) Then OpenFile = picked(LBound(picked))
End Function

' Shows an Open dialog allowing several files; returns the paths, or an empty array when cancelled.
Public Function OpenFiles(Optional ByVal title As String = "", Optional ByVal filterName As String = "", Optional ByVal filterPattern As String = "", Optional ByVal initialDirectory As String = "") As String()
    Validate "DotNetDialog.OpenFiles", filterName, filterPattern, initialDirectory
    OpenFiles = ShowDialog("DotNetDialog.OpenFiles", MSO_FILE_DIALOG_FILE_PICKER, title, filterName, filterPattern, initialDirectory, True)
End Function

' Shows a Save As dialog; returns the chosen path, or "" when cancelled. Not supported in Access.
Public Function SaveFile(Optional ByVal title As String = "", Optional ByVal filterName As String = "", Optional ByVal filterPattern As String = "", Optional ByVal initialDirectory As String = "") As String
    Dim picked() As String
    If Application.Name = "Microsoft Access" Then
        ThrowDotNet DotNetNotSupported, "DotNetDialog.SaveFile", "The Save As dialog is not supported in Microsoft Access."
    End If
    Validate "DotNetDialog.SaveFile", filterName, filterPattern, initialDirectory
    picked = ShowDialog("DotNetDialog.SaveFile", MSO_FILE_DIALOG_SAVE_AS, title, filterName, filterPattern, initialDirectory, False)
    If UBound(picked) >= LBound(picked) Then SaveFile = picked(LBound(picked))
End Function

' Shows a folder picker; returns the chosen folder, or "" when cancelled.
Public Function SelectFolder(Optional ByVal title As String = "", Optional ByVal initialDirectory As String = "") As String
    Dim picked() As String
    Validate "DotNetDialog.SelectFolder", "", "", initialDirectory
    picked = ShowDialog("DotNetDialog.SelectFolder", MSO_FILE_DIALOG_FOLDER_PICKER, title, "", "", initialDirectory, False)
    If UBound(picked) >= LBound(picked) Then SelectFolder = picked(LBound(picked))
End Function

Private Sub Validate(ByVal source As String, ByVal filterName As String, ByVal filterPattern As String, ByVal initialDirectory As String)
    If (Len(filterName) = 0) <> (Len(filterPattern) = 0) Then
        ThrowDotNet DotNetArgument, source, "filterName and filterPattern must both be given or both be empty."
    End If
    If Len(initialDirectory) > 0 Then
        If Not NewFso().FolderExists(initialDirectory) Then
            ThrowDotNet DotNetDirectoryNotFound, source, "Could not find a part of the path '" & initialDirectory & "'."
        End If
    End If
End Sub

' Returns the selected paths; an empty array when cancelled.
Private Function ShowDialog(ByVal source As String, ByVal dialogType As Long, ByVal title As String, ByVal filterName As String, ByVal filterPattern As String, ByVal initialDirectory As String, ByVal multiSelect As Boolean) As String()
    Dim dlg As Object
    Dim result() As String
    Dim idx As Long
    Dim cnt As Long
    On Error GoTo Fail
    Set dlg = Application.FileDialog(dialogType)
    If Len(title) > 0 Then dlg.Title = title
    If Len(initialDirectory) > 0 Then dlg.InitialFileName = initialDirectory & "\"
    If Len(filterName) > 0 Then
        ' Save As dialogs do not support Filters on every host; ignore failures there only.
        If dialogType = MSO_FILE_DIALOG_SAVE_AS Then On Error Resume Next
        dlg.Filters.Clear
        dlg.Filters.Add filterName, filterPattern
        If dialogType = MSO_FILE_DIALOG_SAVE_AS Then
            Err.Clear
            On Error GoTo Fail
        End If
    End If
    dlg.AllowMultiSelect = multiSelect
    If dlg.Show = -1 Then
        cnt = dlg.SelectedItems.Count
        If cnt > 0 Then
            ReDim result(0 To cnt - 1)
            For idx = 1 To cnt
                result(idx - 1) = CStr(dlg.SelectedItems.Item(idx))
            Next idx
            ShowDialog = result
            Exit Function
        End If
    End If
    ShowDialog = EmptyStringArray()
    Exit Function
Fail:
    ThrowDotNet DotNetIO, source, Err.Description
End Function
