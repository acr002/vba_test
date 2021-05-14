Attribute VB_Name = "SetupJournal"
'-----------------------------------------------------------[date: 2021.05.14]
Option Explicit

'***********************************************
' 2018.06.28(��) �\���̉�������E������͊֐���
' 2018.08.20(��) �ꐔ�ύX 1.08 > 1.07
' 2019.04.25(��)
' 2019.06.17(��) �ꐔ�ύX 1.05 > 1.04
' 2019.06.24(��) �ꐔ�ύX 1.04 > 1.03
' 2020.06.08(��) 2020�N�p�ɕύX
' 2020.06.08(��) �ꐔ�ύX 1.03 > 1.02
' 2021.05.12(��) 2021�N�p�ɕύX
'***********************************************

Private Type ar_variant
  ws As Worksheet
  rc As Range
  gc As Variant
  ys As Long
  xs As Long
End Type
'***********************************************
Private Function cc_arv_v2(ws As Worksheet) As ar_variant
  Dim a As ar_variant
  Set a.ws = ws
  Set a.rc = a.ws.UsedRange
  a.gc = a.rc.Value2
  If IsArray(a.gc) Then
    a.ys = UBound(a.gc, 1)
    a.xs = UBound(a.gc, 2)
  Else
    a.ys = 1
    a.xs = 1
  End If
  cc_arv_v2 = a
End Function
'-----------------------------------------------------------------------------

Private Function find_wb() As Workbook
  Dim b_wb As Workbook
  If Workbooks.Count = 2 Then
    For Each b_wb In Workbooks
      If Not b_wb Is ThisWorkbook Then
        Set find_wb = b_wb
        Exit For
      End If
    Next b_wb
  End If
End Function
'-----------------------------------------------------------------------------

Public Sub setup_journal()
  ' Const PARAMETER As Single = 1.09   ': �ꐔ(�����)
  ' Const PARAMETER As Single = 1.11
  ' Const PARAMETER As Single = 1.10
  ' Const PARAMETER As Single = 1.02
  Const VALUES1_SIZE As Long = 50
  Const VALUES2_SIZE As Long = 62
  Const PARAMETER    As Single = 1.01
  Const BASE_DATE    As Long = 44328
  Const X_ERROR      As Long = 4
  Const X_NORMAL     As Long = 5
  Dim ar2()    As Long
  Dim ar1()    As Long
  Dim key      As Long
  Dim t        As Long
  Dim y        As Long
  Dim a        As ar_variant
  Dim wb       As Workbook
  Dim ws       As Worksheet
  Dim cn_entry As Long
  Dim cn_error As Long
  Set wb = find_wb()
  If wb Is Nothing Then
    MsgBox "�Ώۃt�@�C�����J���Ă�������"
    Exit Sub
  End If
  a = cc_arv_v2(wb.Worksheets("list"))
  ReDim ar1(VALUES1_SIZE - 1)
  ReDim ar2(VALUES2_SIZE - 1)
  For y = 2 To a.ys
    key = Val(a.gc(y, X_NORMAL))
    If key Then
      cn_entry = cn_entry + 1
      t = key - BASE_DATE
      Select Case t
        Case Is < VALUES1_SIZE
          ar1(t) = ar1(t) + 1
        Case Else
          t = t - VALUES1_SIZE
          ar2(t) = ar2(t) + 1
      End Select
    Else
      If Len(a.gc(y, X_ERROR)) Then
        cn_error = cn_error + 1
      End If
    End If
  Next y
  Set ws = wb.Worksheets("journal")
  ws.Cells(7, 3).Value = Now
  ws.Cells(10, 3).Value = cn_error
  ws.ChartObjects(1).Chart.SeriesCollection(1).Values = cn_entry / PARAMETER
  ws.ChartObjects(1).Chart.SeriesCollection(2).Values = 100 - (cn_entry / PARAMETER)
  ws.ChartObjects(2).Chart.SeriesCollection(1).Values = ar1
  ws.ChartObjects(3).Chart.SeriesCollection(1).Values = ar2
End Sub
'-----------------------------------------------------------------------------


