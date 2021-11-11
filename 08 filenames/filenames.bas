'-----------------------------------------------------------[date: 2021.11.11]
Attribute VB_Name = "Module1"
Option Explicit

' 再帰的にファイル名を取得したコレクションを返します。
' 対象のフォルダを引数に渡してください。
' 配布用
Private Function get_filenames_sub(ByVal a_path As String) As Collection
  Dim fso      As Object
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim b_file   As Object
  Dim b_folder As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set cc = New Collection
  For Each b_file In fso.getfolder(a_path).Files
    cc.Add b_file.path
  Next b_file
  For Each b_folder In fso.getfolder(a_path).subfolders
    Set r_cc = get_filenames_sub(b_folder.path)
    For Each ii In r_cc
      cc.Add ii
    Next ii
  Next b_folder
  Set fso = Nothing
  Set get_filenames_sub = cc
End Function
'-----------------------------------------------------------------------------

Public Sub main_filenames()
  Dim fns  As Collection
  Dim ii   As Variant
  Dim path As String
  path = ThisWorkbook.path & "\in\"
  Set fns = get_filenames_sub(path)
  For Each ii In fns
    Debug.Print ii
  Next ii
End Sub
'-----------------------------------------------------------------------------

Public Sub main_filenames_dialog()
  Dim path_this As String
  Dim path      As String
  Dim fns       As Collection
  Dim ii        As Variant
  path_this = ThisWorkbook.path & "\"
  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = path_this
    .InitialView = msoFileDialogViewDetails
    .Title = "『フォルダ』を選んでください"
    If .Show = True Then
      path = .SelectedItems(1)
    Else
      path = path_this
    End If
  End With
  Set fns = get_filenames_sub(path)
  For Each ii In fns
    Debug.Print ii
  Next ii
End Sub
'-----------------------------------------------------------------------------


