'-----------------------------------------------------------[date: 2021.06.17]
Attribute VB_Name = "CDmain"
Option Explicit

Private Function get_filenames_sub(ByVal a_path As String) As Collection
  Dim fso      As Object
  Dim r_cc     As Collection
  Dim cc       As Collection
  Dim ii       As Variant
  Dim b_file   As Object
  Dim b_folder As Object
  Set fso = Createobject("Scripting.FileSystemObject")
  Set cc = New Collection
  For Each b_file In fso.getfolder(a_path).files
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

Public Sub main()
  Dim path As String
  Dim a_xs As Long
  Dim a_ys As Long
  Dim py   As Long
  Dim a    As Variant
  Dim ws   As Worksheet
  Dim wb   As Workbook
  Dim pws  As Worksheet
  Dim pwb  As Workbook
  Dim fn   As Variant
  Dim fns  As Collection
  Set fns = get_filenames_sub(ThisWorkbook.path & "\in")
  Set pwb = Workbooks.Add
  Set pws = pwb.Worksheets(1)
  py = 1
  For Each fn In fns
    If fn Like "*.xls*" Then
      Set wb = Workbooks.Open(Filename:=fn, ReadOnly:=True)
      For Each ws In wb.Worksheets
        a = ws.UsedRange.Value
        If IsEmpty(a) Then
        Else
          If IsArray(a) Then
            a_ys = UBound(a, 1)
            a_xs = UBound(a, 2)
            pws.Cells(py, 1).Resize(a_ys, a_xs) = a
            py = py + a_ys
          Else
            pws.Cells(py, 1).Value = a
            py = py + 1
          End If
        End If
      Next ws
      wb.Close savechanges:=False
    End If
  Next fn
  path = ThisWorkbook.path & "\result " & Format(Now, "yyyymmddhhmmss")
  pwb.SaveAs Filename:=path
  pwb.Close
  MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

