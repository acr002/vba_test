'-----------------------------------------------------------[date: 2021.07.15]
Attribute VB_Name = "Module1"
Option Explicit

' 2021.07.15(ñÿ)

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

Public Sub main_table_interior()
  dim xs as long
  Dim x As Long
  Dim y As Long
  Dim a As ar_variant
  a = cc_arv(ActiveSheet)
  For y = 1 To a.ys
    If Len(Trim(a.gc(y, 2))) Then
      If a.gc(y, 2) <> "ï\ëË" Then
        for x = 4 to a.xs
          if len(a.gc(y, x)) = 0 then
            xs = x - 1
            exit for
          end if
        next x
        if xs = 0 then
          xs = a.xs
        end if
        If Len(Trim(a.gc(y, 3))) Then
          a.ws.cells(y, 4).resize(1, xs).NumberFormatLocal = "#,0""åè"""
        Else
          a.ws.cells(y, 4).resize(1, xs).NumberFormatLocal = "0.0""%"""
        End If
        xs = 0
      End If
    End If
  Next y
  MsgBox "end of run"
End Sub
'-----------------------------------------------------------------------------

