'-----------------------------------------------------------[date: 2021.08.27]
Attribute VB_Name = "Module1"
Option Explicit

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

Private Function basename_ext(path As String) As String
  Dim buf     As String
  Dim t_fn    As String
  Dim i       As Long
  Dim ar      As Variant
  Dim ar_last As Long
  ar = Split(path, "\")
  ar_last = UBound(ar)
  If Len(Trim(ar(ar_last))) Then
    t_fn = ar(ar_last)
  Else
    t_fn = ar(ar_last - 1)
  End If
  ar = Split(t_fn, ".")
  For i = LBound(ar) To UBound(ar) - 1
    If i = LBound(ar) Then
      buf = ar(i)
    Else
      buf = buf & "." & ar(i)
    End If
  Next i
  basename_ext = buf
End Function
'-----------------------------------------------------------------------------

Public Sub main_to_pdf()
  Dim put_fn As String
  Dim ws As Worksheet
  Dim wb As Workbook
  Dim fn As Variant
  Dim fns As Collection
  Dim this_path As String
  this_path = ThisWorkbook.path & "\"
  Set fns = get_filenames_sub(this_path & "in")
  For Each fn In fns
    If fn Like "*xls*" Then
      Set wb = Workbooks.Open(Filename:=fn, ReadOnly:=True)
      For Each ws In wb.Worksheets
        if ws.visible = true then
          put_fn = this_path & basename_ext(CStr(fn)) & Space(1) & ws.Name
          ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=put_fn
        end if
      Next ws
      wb.Close savechanges:=False
    End If
  Next fn
  MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

Public Sub main_to_pdf_book()
  Dim put_fn As String
  Dim wb As Workbook
  Dim fn As Variant
  Dim fns As Collection
  Dim this_path As String
  this_path = ThisWorkbook.path & "\"
  Set fns = get_filenames_sub(this_path & "in")
  For Each fn In fns
    If fn Like "*xls*" Then
      Set wb = Workbooks.Open(Filename:=fn, ReadOnly:=True)
      put_fn = this_path & basename_ext(CStr(fn))
      wb.ExportAsFixedFormat Type:=xlTypePDF, Filename:=put_fn
      wb.close savechanges:=false
    end if
  Next fn
  MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

