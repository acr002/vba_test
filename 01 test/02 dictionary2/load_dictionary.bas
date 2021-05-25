'-----------------------------------------------------------[date: 2021.05.25]
Attribute VB_Name = "Module1"
Option Explicit

'###############################
' 2021.05.25(‰Î)
'###############################
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

Private Function load_dictionary(ws As Worksheet, Optional xk As Long = 2, Optional xv As Long = 3) As Variant
  Dim y As Long
  Dim dic As Variant
  Dim ys As Long
  Dim a As Variant
  a = ws.UsedRange.Value
  ys = UBound(a, 1)
  Set dic = CreateObject("scripting.dictionary")
  For y = 2 To ys
    If Len(Trim(CStr(a(y, xk)))) <> 0 Then
      dic(a(y, xk)) = a(y, xv)
    End If
  Next y
  Set load_dictionary = dic
End Function
'-----------------------------------------------------------------------------

Private Function dic_value(dic As Variant, key As Variant, Optional none As String = "") As Variant
  If dic.exists(key) Then
    dic_value = dic(key)
  Else
    dic_value = none
  End If
End Function
'-----------------------------------------------------------------------------

Public Sub test01()
  dim a as ar_variant
  Dim y As Long
  Dim bb As Variant
  Dim dic As Variant
  Set dic = load_dictionary(ThisWorkbook.Worksheets("list"))
  a = cc_arv(thisworkbook.worksheets("work"))
  For y = 2 To a.ys
    a.ws.Cells(y, 3).Value = dic_value(dic, a.gc(y, 2), "nothing")
  Next y
End Sub
'-----------------------------------------------------------------------------

