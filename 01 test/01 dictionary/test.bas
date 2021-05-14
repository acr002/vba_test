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



