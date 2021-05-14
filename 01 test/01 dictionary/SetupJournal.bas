Attribute VB_Name = "SetupJournal"
'-----------------------------------------------------------[date: 2021.05.14]
Option Explicit

'***********************************************
' 2018.06.28(木) 表紙の回収件数・回収率は関数化
' 2018.08.20(月) 母数変更 1.08 > 1.07
' 2019.04.25(木)
' 2019.06.17(月) 母数変更 1.05 > 1.04
' 2019.06.24(月) 母数変更 1.04 > 1.03
' 2020.06.08(月) 2020年用に変更
' 2020.06.08(月) 母数変更 1.03 > 1.02
' 2021.05.12(水) 2021年用に変更
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

Public Sub setup_journal()
  ' Const PARAMETER As Single = 1.09   ': 母数(会員数)
  ' Const PARAMETER As Single = 1.11
  ' Const PARAMETER As Single = 1.10
  ' Const PARAMETER As Single = 1.02
  Const VALUES1_SIZE As Long = 50
  Const VALUES2_SIZE As Long = 62
  Const PARAMETER    As Single = 1.01
  Const BASE_DATE    As long = 44328
  Const X_ERROR      As Long = 4
  Const X_NORMAL     As Long = 5
  dim bb as variant
  Dim key      As long
  Dim y        As Long
  Dim dic(2)   As Variant
  Dim a        As ar_variant
  Dim wb       As Workbook
  Dim ws       As Worksheet
  Dim cc       As Chart
  Dim cn_entry As Long
  Dim cn_error As Long
  Set wb = find_wb()
  If wb Is Nothing Then
    MsgBox "対象ファイルも開いてください"
    Exit Sub
  End If
  a = cc_arv_v2(wb.Worksheets("list"))
  Set dic(0) = CreateObject("scripting.dictionary")
  Set dic(1) = CreateObject("scripting.dictionary")
  Set dic(2) = CreateObject("scripting.dictionary")
  For y = 0 To VALUES1_SIZE - 1
    dic(1)(y + BASE_DATE) = 0
  Next y
  For y = 0 To VALUES2_SIZE - 1
    dic(2)(y + BASE_DATE + VALUES1_SIZE) = 0
  Next y
  For y = 2 To a.ys
    key = Val(a.gc(y, X_NORMAL))
    If key Then
      cn_entry = cn_entry + 1
      Select Case key
        Case BASE_DATE To (BASE_DATE + VALUES1_SIZE - 1)
          dic(1)(key) = dic(1)(key) + 1
        Case (BASE_DATE + VALUES1_SIZE) To (BASE_DATE + VALUES1_SIZE + VALUES2_SIZE - 1)
          dic(2)(key) = dic(2)(key) + 1
        Case Else
          If dic(0).exists(key) Then
            dic(0)(key) = dic(0)(key) + 1
          Else
            dic(0)(key) = 1
          End If
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
  Set cc = ws.ChartObjects(1).Chart
  cc.SeriesCollection(1).Values = cn_entry / PARAMETER
  cc.SeriesCollection(2).Values = 100 - (cn_entry / PARAMETER)
  ws.ChartObjects(2).Chart.SeriesCollection(1).Values = dic(1).items
  ws.ChartObjects(3).Chart.SeriesCollection(1).Values = dic(2).items
End Sub
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

