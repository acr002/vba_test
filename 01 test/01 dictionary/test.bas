'-----------------------------------------------------------[date: 2021.05.14]
Attribute VB_Name = "Module1"
Option Explicit

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
  if isarray(a.gc) then
    a.ys = UBound(a.gc, 1)
    a.xs = UBound(a.gc, 2)
  else
    a.ys = 1
    a.xs = 1
  end if
  cc_arv = a
End Function
'-----------------------------------------------------------------------------


' https://qiita.com/nna1016/questions/beef01479dca171b1ab5
Sub sample()
  Dim d, a, k, i
  Set d = CreateObject("Scripting.Dictionary")
  a = Split(ActiveCell.Value, vbLf)
  k = a(0)
  d(k) = 0
  For i = 1 To UBound(a)
    If a(i) - a(i - 1) = 1 Then
      d(k) = d(k) + 1
    Else
      k = a(i)
      d(k) = 0
    End If
  Next
  For Each k In d
    d(k) = IIf(d(k) = 0, k, k & "-" & k + d(k))
  Next
  ActiveCell.Offset(, 1).Value = Join(d.Items, vbLf)
End Sub
'-----------------------------------------------------------------------------

