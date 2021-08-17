'-----------------------------------------------------------[date: 2021.08.17]
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

PRINT    FORMAT A2 (F8).
VARIABLE LABELS A2 'D2'.
   VALUE LABELS A2 1  'ct1' 2  'ct2'.


public sub main_spss_syntax()
  const PRINT_FORMAT as string    = "PRINT    FORMAT "
  const VARIABLE_LABELS as string = "VARIABLE LABELS "
  const VALUE_LABELS as string    = "   VALUE LABELS "
  dim buf as string
  dim x as long
  dim y as long
  dim cc as collection
  dim a as ar_variant
  a = cc_arv(thisworkbook.worksheets(1))
  set cc = new collection
  for y = 2 to a.ys
    cc.add PRINT_FORMAT & a.gc(y, 1) & "(" & a.gc(y, 3) & ")"




    if len(a.gc(y, 1)) * len(a.gc(y, 2)) * len(a.gc(y, 3)) then


End sub
'-----------------------------------------------------------------------------

