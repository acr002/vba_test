'-----------------------------------------------------------[date: 2021.07.05]
Attribute VB_Name = "Module1"
Option Explicit

' 2021.07.02(ã‡)

Private Type ar_variant
  ws As Worksheet
  rc As Range
  gc As Variant
  ys As Long
  xs As Long
End Type
'***********************************************
Private Function cc_arv(ws As Worksheet) As ar_variant
  Dim a As ar_variant
  Set a.ws = ws
  Set a.rc = a.ws.UsedRange
  a.gc = a.rc.Value
  If IsArray(a.gc) Then
    a.ys = UBound(a.gc, 1)
    a.xs = UBound(a.gc, 2)
  Else
    a.ys = 1
    a.xs = 1
  End If
  cc_arv = a
End Function
'-----------------------------------------------------------------------------

Public Sub main_entry_check()
  Dim bb  As Variant
  Dim col As Collection
  Dim x   As Long
  Dim y   As Long
  Dim b   As ar_variant
  Dim a   As ar_variant
  a = cc_arv(ThisWorkbook.Worksheets("entry"))
  b = cc_arv(ThisWorkbook.Worksheets("verify"))
  Set col = New Collection
  For y = 2 To a.ys
    For x = 1 To a.xs
      If a.gc(y, x) <> b.gc(y, x) Then
        ' 2, 4, 12, 16çsñ⁄ÇÕëŒè€äOÇ∆ÇµÇ‹Ç∑ÅB
        Select Case x
          Case 2, 4, 12, 16
          Case Else
            col.Add CStr(a.gc(y, 1)) & ": " & CStr(x)
            b.rc(y, x).interior.color = rgb(255, 0, 0)
        End Select
      End If
    Next x
  Next y
  If col.Count Then
    For Each bb In col
      Debug.Print bb
    Next bb
  End If
  MsgBox "end of run" & vbCrLf & "Error: " & CStr(col.Count)
End Sub
'-----------------------------------------------------------------------------


