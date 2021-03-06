Attribute VB_Name = "Module00"
Option Explicit

Public incol_array() As String
Public wb            As Workbook
Public wb_indata     As Workbook
Public ws_indata     As Worksheet
Public wb_outdata    As Workbook
Public ws_outdata    As Worksheet
Public ws_mainmenu   As Worksheet
Public ws_setup      As Worksheet

Public gcode_row, gcode_col     As Integer
Public gdrive_row, gdrive_col   As Integer
Public initial_row, initial_col As Integer
Public setup_row, setup_col     As Integer
Public incol_row, incol_col     As Integer

Public file_path    As String
Public ope_code     As String
Public outdata_fn   As String
Public status_msg   As String
Public progress_msg As String

' 20170419 ºR½ õpQCODEzñ
Public str_code() As String

' 20170425 ºR½ ÝèæÊi[p[U^
Public Type question_data
q_code As String                ' QCODEi[p¶ñÏ
r_code As String                ' ÀCODEi[p¶ñÏ
m_code As String                ' MCODEi[p¶ñÏ
q_format As String              ' Ýâ`®i[p¶ñÏ
r_byte As Integer               ' Ài[pÏ
r_unit As String                ' ÀPÊ¶ñÏi2017.05.10 ÇÁj
sel_code1 As String             ' ZNgð@QCODEi[p¶ñÏ
sel_value1 As Integer           ' ZNgð@li[p¶ñÏ
sel_code2 As String             ' ZNgðAQCODEi[p¶ñÏ
sel_value2 As Integer           ' ZNgðAli[p¶ñÏ
sel_code3 As String             ' ZNgðBQCODEi[p¶ñÏ
sel_value3 As Integer           ' ZNgðBli[p¶ñÏ
ct_count As Integer             ' ÝâJeS[i[pÏ
ct_loop As Integer              ' [vJEgi[pÏ
'        ct_0flg As Boolean              ' OJeS[tOi[pÏ
q_title As String               ' \èi[p¶ñÏ
q_ct(300) As String             ' Ýâi[p¶ñzñ
data_column As Integer          ' üÍf[^ñÔi[pÏ
End Type

Public q_data() As question_data    ' QCODEÌf[^ðSÄæ¾

public Sub Auto_Open()
  Dim base_pt As String
  '--------------------------------------------------------------------------------------------------'
  '@t@CI[v@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.10@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2020.06.03  '
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  '    Call Starting_Mcs2017    ' ThisWorkbook ÉÄÄÑoµÉÏX - 2020.6.3
  Application.StatusBar = "t@CI[v..."
  Application.ScreenUpdating = False
  wb.Activate
  ws_mainmenu.Select
  If ws_mainmenu.Cells(gdrive_row, gdrive_col) <> "" And _
    ws_mainmenu.Cells(gcode_row, gcode_col) <> "" Then
    base_pt = ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If Dir(base_pt & "\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
      If (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 8) <> "// ÇÝñ¾ú") And _
        (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 7) <> "// Û¶µ½ú") Then
        ws_mainmenu.Cells(initial_row, initial_col) = "// úÝèÏÝ"
      End If
    Else
      ws_mainmenu.Cells(initial_row, initial_col) = ""
    End If
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
  End If
  Application.ScreenUpdating = True
  Application.StatusBar = False
End Sub
'-----------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------'
'@V[g\¬ÌúÝè@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
'--------------------------------------------------------------------------------------------------'
'@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.10@'
'@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2017.04.14@'
'--------------------------------------------------------------------------------------------------'
public Sub Starting_Mcs2017()
  Set wb = ThisWorkbook
  Set ws_mainmenu = wb.Worksheets("Cj[")
  Set ws_setup = wb.Worksheets("ÝèæÊ")
  ' Cj[FÆ±R[hÌsñ
  gcode_row = 3
  gcode_col = 8
  ' Cj[FìÆhCuÌsñ
  gdrive_row = 3
  gdrive_col = 23
  ' Cj[FúÝèÏÝbZ[WoÍæsñ
  initial_row = 6
  initial_col = 32
  ' ÝèæÊFæªp[^Ìsñ
  setup_row = 3
  setup_col = 1
  ' ÝèæÊFüÍf[^ñÔÌsñ
  incol_row = 3
  incol_col = 3
End Sub
'-----------------------------------------------------------------------------

public Sub Finishing_Mcs2017()
  '--------------------------------------------------------------------------------------------------'
  '@I¹EeIuWFNgÌQÆðð@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.10@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2018.09.19@'
  '--------------------------------------------------------------------------------------------------'
  Application.ScreenUpdating = True
  Application.StatusBar = False

  wb.Activate
  ws_setup.Select
  ws_setup.Cells(3, 1).Select
  ws_mainmenu.Select
  ws_mainmenu.Cells(3, 8).Select

  ' ÝèæÊ`FbNÌG[Ot@Cª0oCgÈçt@Cí
  If Dir(file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx") <> "" Then
    If FileLen(file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx") = 0 Then
      Kill file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx"
    End If
  End If

  ' WbN`FbNÉæéf[^t@CÌG[Ot@Cª0oCgÈçt@Cí
  If Dir(file_path & "\4_LOG\" & ope_code & "err.xlsx") <> "" Then
    If FileLen(file_path & "\4_LOG\" & ope_code & "err.xlsx") = 0 Then
      Kill file_path & "\4_LOG\" & ope_code & "err.xlsx"
    End If
  End If

  Application.DisplayAlerts = False
  wb.Save
  Application.DisplayAlerts = True

  Set wb = Nothing
  Set ws_mainmenu = Nothing
  Set ws_setup = Nothing
End Sub
'-----------------------------------------------------------------------------

public Sub Setup_Check()
  Dim qcode_cnt As Long
  Dim setup_status As String

  Dim wb_preset As Workbook
  Dim ws_preset As Worksheet
  Dim preset_row As Variant
  Dim preset_col As Variant
  Dim preset_i As Long
  Dim preset_j As Long
  Dim preset_gcode As String

  Dim max_row As Long
  Dim max_col As Long
  Dim err_msg As String
  Dim err_cnt As Long
  Dim sel_i As Long
  Dim selcode_row As Long
  Dim sys_cnt As Long
  Dim sys_num As Long
  Dim sys_col As Long  '5,6,7
  Dim sys_i As Long
  Dim mcode_col As Long
  Dim qCStr_col As Long
  Dim qformat_typ As String
  Dim selcode_col_1 As Long
  Dim selval_col_1 As Long
  Dim selcode_col_2 As Long
  Dim selval_col_2 As Long
  Dim selcode_col_3 As Long
  Dim selval_col_3 As Long
  Dim find_row_1 As Long
  Dim find_row_2 As Long
  Dim find_row_3 As Long
  Dim ct_cnt_col As Long
  Dim zero_f_col As Long
  Dim qttl_col As Long
  Dim ct_st_col As Long
  Dim rcode_col As Long
  Dim rcode_row As Long
  Dim rcode_cell As Range
  '--------------------------------------------------------------------------------------------------'
  '@ÝèæÊ`FbN@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.12@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2020.05.19@'
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  Call Starting_Mcs2017
  Application.StatusBar = "ÝèæÊ `FbN..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If ws_mainmenu.Cells(gcode_row, gcode_col).Value = "" Then
    MsgBox "Cj[ÌÆ±R[hª¢üÍÅ·B", vbExclamation, "MCS 2020 - Setup_Check"
    Cells(gcode_row, gcode_col).Select
    Application.StatusBar = False
    End
  End If

  Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & _
  "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & "_mcs.ini" For Input As #1
  Line Input #1, file_path
  Line Input #1, setup_status
  Close #1

  Open file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx" For Append As #1
  Close #1
  If Err.Number > 0 Then
    Workbooks(ope_code & "_ÝèæÊNG.xlsx").Close
  End If

  If Dir(file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx") <> "" Then
    Kill file_path & "\4_LOG\" & ope_code & "_ÝèæÊNG.xlsx"
  End If

  ws_setup.Select
  max_row = Cells(Rows.Count, setup_col).End(xlUp).Row

  ' ±±©çG[`FbNÌR[fBO(L¥Ö¥`)
  err_msg = "yÝèæÊÌG[Xgz" & vbCrLf & "QCODE,s,G[Ú,G[àe" & vbCrLf

  'Tvio[ÌQCODE`FbN
  If Cells(3, 1).Value <> "SNO" Then
    err_msg = err_msg & Cells(3, 1).Value & ",3sÚ" & ",QCODE(RC1)" & ",TvÌQCODEÍmSNOnÆµÄ­¾³¢B" & vbCrLf
  End If

  For qcode_cnt = setup_row To max_row
    If Cells(qcode_cnt, setup_col).Value = "*ÁHã" Then
      Rows(qcode_cnt).Select
      With Selection.Interior
        .Color = 65535
      End With
    ElseIf Left(Cells(qcode_cnt, setup_col).Value, 1) = "*" Then
      '³µ
    Else
      'ÝèæÊsñîño^
      max_col = Cells(qcode_cnt, Columns.Count).End(xlToLeft).Column
      sys_cnt = 3: sys_num = 5
      mcode_col = 8: qCStr_col = 9
      selcode_col_1 = 10: selval_col_1 = 11
      selcode_col_2 = 12: selval_col_2 = 13
      selcode_col_3 = 14: selval_col_3 = 15
      ct_cnt_col = 16: zero_f_col = 17
      qttl_col = 18: ct_st_col = 19
      find_row_1 = 0
      find_row_2 = 0
      find_row_3 = 0
      rcode_col = 2: rcode_row = 0

    'QCODE Ìd¡`FbN
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
    Cells(qcode_cnt, setup_col).Value)
    If err_cnt >= 2 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",QCODE(RC1)" & ",QCODEÍd¡sÂÅ·B" & vbCrLf
    End If

    'QCODE¢wè©ÂñRgsÖÌüÍ`FbN
    If Cells(qcode_cnt, setup_col).Value = "" And _
      WorksheetFunction.CountA(Range(Cells(qcode_cnt, setup_col), Cells(qcode_cnt, max_col))) <> 0 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",QCODE(RC1)" & ",QCODEªuNÌêA¼ÌÝèÚÖÌüÍÍsÂÆÈèÜ·B" & vbCrLf
    End If

    'MCODEPÆüÍ`FbN
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, mcode_col), Cells(max_row, mcode_col)), _
    Cells(qcode_cnt, mcode_col).Value)
    If err_cnt = 1 And Cells(qcode_cnt, mcode_col).Value <> "" Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",MCODE(RC8)" & ",MCODEÍ2ÂÈãÌÝâÉÎµÄwèµÄ­¾³¢B" & vbCrLf
    End If

    '`®`FbN
    qformat_typ = Cells(qcode_cnt, qCStr_col).Value
    Select Case Left(qformat_typ, 1)
      Case "C", "S", "M", "H", "F", "O"
        If Len(qformat_typ) > 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",`®Í1ÅmCnmSnmMnmHnmFnmOnðwèµÄ­¾³¢B" & vbCrLf
        End If
      Case "R"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",`®mRnÍðwèµÄ­¾³¢B" & vbCrLf
        ElseIf Len(qformat_typ) > 3 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",Í2ÜÅwèÂ\Å·B" & vbCrLf
        ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",Í¼pÅwèµÄ­¾³¢B" & vbCrLf
        End If
      Case "L"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",`® mLnÍñ§ÀÌwèðµÄ­¾³¢B" & vbCrLf
        ElseIf (Mid(qformat_typ, 2, 1) = "M") Or (Mid(qformat_typ, 2, 1) = "A") Or (Mid(qformat_typ, 2, 1) = "C") Then
          If Len(qformat_typ) = 2 Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",`® mLnÍñ§ÀÌwèðµÄ­¾³¢B" & vbCrLf
          ElseIf Len(qformat_typ) > 5 And IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",ñ§ÀÍ3ÜÅwèÂ\Å·B" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",ñ§ÀÍ¼pÅwèµÄ­¾³¢B" & vbCrLf
          End If
        ElseIf Val(Mid(qformat_typ, 2, 1)) = 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",~bg}`Ì`®ÍmLMxnmLAxnmLCxnÌ¢¸ê©ÅwèðµÄ­¾³¢ixÍñjB" & vbCrLf
        Else
          If Len(qformat_typ) > 4 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",ñ§ÀÍ3ÜÅwèÂ\Å·B" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",ñ§ÀÍ¼pÅwèµÄ­¾³¢B" & vbCrLf
          End If
        End If
      Case Else
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",`®(RC9)" & ",`®mCnmSnmMnmHnmFnmOnmRnmLnÈOÌwèÍsÂÆÈèÜ·B" & vbCrLf
    End Select

    'ZNg`FbN
    err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3)))
    If err_cnt <> 0 Then
      'ZNgðL@`FbN
      If WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_1))) < 2 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_2), Cells(qcode_cnt, selval_col_2))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@AB(RC10:RC15)" & ",ðiZNgjÌwèÍQCODEÆlÌZbgðð@©ç¶lßÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_2))) < 4 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@AB(RC10:RC15)" & ",ðiZNgjÌwèÍQCODEÆlÌZbgðð@©ç¶lßÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3))) < 6 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@AB(RC10:RC15)" & ",ðiZNgjÌwèÍQCODEÆlÌZbgðð@©ç¶lßÅwèµÄ­¾³¢B" & vbCrLf
      End If

      'ZNgðQCODEY`FbN
      'ðP
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_1).Value) = 0 And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ðiZNgjÆÈéQCODEÍ\ßAñÉwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        find_row_1 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_1).Value, lookat:=xlWhole).Row
      End If
      'ð2
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_2).Value) = 0 And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðiZNgjÆÈéQCODEÍ\ßAñÉwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        find_row_2 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_2).Value, lookat:=xlWhole).Row
      End If
      'ð3
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_3).Value) = 0 And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðiZNgjÆÈéQCODEÍ\ßAñÉwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        find_row_3 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_3).Value, lookat:=xlWhole).Row
      End If

      'ZNgðl`FbN
      'ðP
      If Cells(qcode_cnt, selval_col_1).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ðiZNgjÌlÍ¼pÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value = "" And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value <> "" And Cells(qcode_cnt, selcode_col_1).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = True And find_row_1 <> 0 Then
        If Cells(find_row_1, ct_cnt_col).Value < Val(Cells(qcode_cnt, selval_col_1).Value) Then
          Debug.Print find_row_1
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ðiZNgjÌlÍYÝâÌIðÈºÌlÅwèµÄ­¾³¢B" & vbCrLf
        End If
      End If
      'ðQ
      If Cells(qcode_cnt, selval_col_2).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðiZNgjÌlÍ¼pÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value = "" And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value <> "" And Cells(qcode_cnt, selcode_col_2).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = True And find_row_2 <> 0 Then
        If Cells(find_row_2, ct_cnt_col).Value - Cells(find_row_2, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_2).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðiZNgjÌlÍYÝâÌIðÈºÌlÅwèµÄ­¾³¢B" & vbCrLf
        End If
      End If
      'ðR
      If Cells(qcode_cnt, selval_col_3).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðiZNgjÌlÍ¼pÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value = "" And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value <> "" And Cells(qcode_cnt, selcode_col_3).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðiZNgjÍQCODEÆlÌZbgÅwèµÄ­¾³¢B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = True And find_row_3 <> 0 Then
        If Cells(find_row_3, ct_cnt_col).Value - Cells(find_row_3, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_3).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðiZNgjÌlÍYÝâÌIðÈºÌlÅwèµÄ­¾³¢B" & vbCrLf
        End If
      End If
    End If

    'Ið`FbN
    Select Case Left(Cells(qcode_cnt, 9).Value, 1)
      Case "S", "M", "L"
        If max_col >= ct_st_col Then
          err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, ct_st_col), Cells(qcode_cnt, max_col)))
          If err_cnt <> Cells(qcode_cnt, ct_cnt_col).Value Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",CTEJeS[àe(RC16RC19`)" & ",PñÌCTÆASñÈ~ÌJeS[ÚÌÂªêvµÜ¹ñB" & vbCrLf
          ElseIf err_cnt <> max_col - qttl_col Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",JeS[Ú(RC19`)" & ",SñÈ~ÌJeS[ÚÍ¶lßÅüÍµÄ­¾³¢B" & vbCrLf
          End If
        End If
      Case "F", "O", "R", "H", "C"
        If max_col >= ct_st_col Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",JeS[Ú(RC19`)" & ",`®mRnmHnmFnmOnmCnÌJeS[ÚÌÝèÍÅ«Ü¹ñB" & vbCrLf
        ElseIf Cells(qcode_cnt, ct_cnt_col).Value <> 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",CT(RC16)" & ",`®mRnmHnmFnmOnmCnÌÝâÌCTÍmuNnÉµÄ­¾³¢B" & vbCrLf
        End If
      Case Else
    End Select

    ' 2020.1.9 - [tOÍÈ­ÈèÜµ½ÌÅARgAEgµÜµ½B
    '            '[tO`FbN
    '            If Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Len(Cells(qcode_cnt, ct_cnt_col).Value) = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",0CT(RC17)" & ",0CTÌtOÍCTªm1ÈãnÌÝâÉÌÝÝèÂ\Å·B" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, ct_cnt_col).Value = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",0CT(RC17)" & ",0CTÌtOÍCTªm1ÈãnÌÝâÉÌÝÝèÂ\Å·B" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, zero_f_col).Value <> 1 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",0CT(RC17)" & ",0CTÌwèÆµÄg¦éÌÍm1nÌÝÆÈèÜ·B" & vbCrLf
    '            End If

    'Ýâ^CgL³`FbN
    If Cells(qcode_cnt, qttl_col) = "" Then
      If Cells(qcode_cnt, ct_st_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",\è(RC18)" & ",ÝâÌ\èÍK¸wèµÄ­¾³¢B" & vbCrLf
      End If
    End If

    '\õGAÖÌüÍ`FbN
    For sys_i = 1 To sys_cnt
      sys_col = sys_num + sys_i - 1
      If Cells(qcode_cnt, sys_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",System(RC5:RC7)" & ",»ÝÌo[WÅÍE~GñÍgpµÈ¢Å­¾³¢B" & vbCrLf
      End If
    Next sys_i

    'ÀwèÝâÆYÀÝâÌZNgóµÌ¯ê`FbN
    If Cells(qcode_cnt, rcode_col).Value <> "" Then
      Set rcode_cell = ws_setup.Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
      Find(What:=Cells(qcode_cnt, rcode_col).Value)
      If rcode_cell Is Nothing Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",À(RC2)" & ",Àwè³ê½QCODEª©Â©èÜ¹ñB" & vbCrLf
      Else
        rcode_row = rcode_cell.Row
        If Cells(qcode_cnt, selcode_col_1).Value <> Cells(rcode_row, selcode_col_1).Value Or _
          Cells(qcode_cnt, selval_col_1).Value <> Cells(rcode_row, selval_col_1).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ð@(RC10:RC11)" & ",ð@iZNgjÌwèªYÌÀÝâÌQCODEm" _
          & Cells(rcode_row, setup_col).Value & "nÆêvµÜ¹ñB" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_2).Value <> Cells(rcode_row, selcode_col_2).Value Or _
          Cells(qcode_cnt, selval_col_2).Value <> Cells(rcode_row, selval_col_2).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðA(RC12:RC13)" & ",ðAiZNgjÌwèªYÌÀÝâÌQCODEm" _
          & Cells(rcode_row, setup_col).Value & "nÆêvµÜ¹ñB" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_3).Value <> Cells(rcode_row, selcode_col_3).Value Or _
          Cells(qcode_cnt, selval_col_3).Value <> Cells(rcode_row, selval_col_3).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "sÚ" & ",ðB(RC14:RC15)" & ",ðBiZNgjÌwèªYÌÀÝâÌQCODEm" _
          & Cells(rcode_row, setup_col).Value & "nÆêvµÜ¹ñB" & vbCrLf
        End If
      End If
    End If
  End If
Next qcode_cnt

Cells(setup_row, setup_col).Select

Application.ScreenUpdating = True
preset_gcode = ws_mainmenu.Cells(gcode_row, gcode_col).Value
If err_msg <> "yÝèæÊÌG[Xgz" & vbCrLf & "QCODE,s,G[Ú,G[àe" & vbCrLf Then
  Application.DisplayAlerts = False
  ' G[bZ[WoÍt@CÌì¬
  Workbooks.Add
  Set wb_preset = ActiveWorkbook
  Set ws_preset = wb_preset.Worksheets("Sheet1")
  Columns("A").NumberFormat = "@"
  If Dir(file_path & "\4_LOG\" & preset_gcode & "_ÝèæÊNG.xlsx") = "" Then
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_ÝèæÊNG.xlsx"
  Else
    Kill file_path & "\4_LOG\" & preset_gcode & "_ÝèæÊNG.xlsx"
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_ÝèæÊNG.xlsx"
  End If
  preset_row = Split(err_msg, vbCrLf)
  For preset_i = 0 To UBound(preset_row)
    preset_col = Split(preset_row(preset_i), ",")
    For preset_j = 0 To UBound(preset_col)
      ws_preset.Cells(preset_i + 1, preset_j + 1).Value = preset_col(preset_j)
    Next preset_j
  Next preset_i
  ws_preset.Activate
  ws_preset.Cells.Select
  With Selection.Font
    .Name = "TakaoSVbN"
    .Size = 11
  End With
  Columns("B:D").EntireColumn.AutoFit
  ws_preset.Cells(1, 1).Select
  wb_preset.Save
  MsgBox "ÝèæÊÌàeÉG[ª èÜ·B" & vbCrLf & _
  file_path & "\4_LOG\" & preset_gcode & "_ÝèæÊNG.xlsx ðmFµÄ­¾³¢B", vbExclamation, "MCS 2020 - Setup_Check"
  Application.DisplayAlerts = True
  ws_preset.Activate
  Application.StatusBar = False
End
  End If
  Application.StatusBar = False
  Set ws_preset = Nothing
  Set wb_preset = Nothing
End Sub

Sub Filepath_Get()
  '--------------------------------------------------------------------------------------------------'
  '@t@CpXÌæ¾@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.10@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2017.05.18@'
  '--------------------------------------------------------------------------------------------------'
  Application.StatusBar = "t@CpXæ¾E¶¬..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
    "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
    "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Input As #1
    Line Input #1, file_path
    Close #1
    ope_code = Cells(gcode_row, gcode_col)
  Else
    MsgBox "Ýèt@Cm" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ininª©Â©èÜ¹ñB" _
    & vbCrLf & "úÝèðà¤êxsÁÄ­¾³¢B", vbExclamation, "MCS 2020 - Filepath_Get"
    End
  End If
  Application.ScreenUpdating = True
  Application.StatusBar = False
End Sub

Sub Indata_Open()
  Dim indata_fn, revdata_fn As String
  '--------------------------------------------------------------------------------------------------'
  '@üÍf[^ÌI[v        @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.11@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2020.04.08@'
  '--------------------------------------------------------------------------------------------------'
  'yTvz@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '@±ÌvV[WÅI[v³êét@CÍAÈºÌp^[Ìt@CÌÝ@@@@@@@@@@@@'
  '@EüÍf[^t@C ...... *IN.xlsx@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '@EC³ãf[^t@C .... *RE.xlsx                  @@@@@@@@@@@@@@@@@@@@@'
  '@ECall³FüÍf[^ÌC³mModule04n                                                          '
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "üÍf[^ I[v..."
  '    Application.ScreenUpdating = False    ' t@CI[vÌi»ª©çê½ûªæ¢ÌÅRgAEg

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & Cells(gcode_row, gcode_col) & "RE.xlsx") <> "" Then
    revdata_fn = Cells(gcode_row, gcode_col) & "RE.xlsx"
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & revdata_fn
    Set wb_indata = Workbooks(revdata_fn)
    Set ws_indata = wb_indata.Worksheets("Sheet1")
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
  ElseIf Dir(Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & Cells(gcode_row, gcode_col) & "IN.xlsx") <> "" Then
    indata_fn = Cells(gcode_row, gcode_col) & "IN.xlsx"
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & indata_fn
    Set wb_indata = Workbooks(indata_fn)
    Set ws_indata = wb_indata.Worksheets("Sheet1")
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
  Else
    MsgBox "üÍf[^t@Cª¶ÝµÜ¹ñB", vbExclamation, "MCS 2020 - Indata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Outdata_Open()
  '--------------------------------------------------------------------------------------------------'
  '@ÁHãf[^ÌI[v@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.14@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2020.04.08@'
  '--------------------------------------------------------------------------------------------------'
  'yTvz@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '@±ÌvV[WÅI[v³êét@CÍAWvÝèt@CÅwè³êÄ¢ét@C@@@@@@'
  '@Call³FWvT}[f[^Ìì¬mModule52n@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "ÁHãf[^ I[v..."
  '    Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn) <> "" Then
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn
    Set wb_outdata = ActiveWorkbook
    Set ws_outdata = wb_outdata.Worksheets("Sheet1")
    Set wb_indata = Nothing
    Set ws_indata = Nothing
  Else
    MsgBox "ÁHãf[^t@Cm" & outdata_fn & "nª¶ÝµÜ¹ñB", vbExclamation, "MCS 2020 - Outdata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Datafile_Open()
  '--------------------------------------------------------------------------------------------------'
  '@f[^t@CÌI[v@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.27@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2020.04.08@'
  '--------------------------------------------------------------------------------------------------'
  'yTvz@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '@±ÌvV[WÅI[v³êét@CÍA[U[©çüÍ³ê½t@C@@@@@@@@@@@'
  '@Call³FüÍf[^ÌWbN`FbNmModule05n@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = status_msg

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn) <> "" Then
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn
    Set wb_outdata = ActiveWorkbook
    Set ws_outdata = wb_outdata.Worksheets("Sheet1")
    Set wb_indata = Nothing
    Set ws_indata = Nothing
  Else
    MsgBox "wèµ½t@Cm" & outdata_fn & "nª¶ÝµÜ¹ñB", vbExclamation, "MCS 2020 - Datafile_Open"
  End
End If

Call Datacol_Get

Application.StatusBar = False
End Sub

Sub Datacol_Get()
  Dim wb_data As Workbook
  Dim ws_data As Worksheet

  Dim max_row, max_col As Long
  Dim i_cnt, setup_cnt As Long
  Dim match_flg As Long
  Dim pass_flg As Integer
  '--------------------------------------------------------------------------------------------------'
  '@f[^t@CÌñÔðæ¾@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.27@'
  '@ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2018.04.04@'
  '--------------------------------------------------------------------------------------------------'
  Application.ScreenUpdating = False

  If Not wb_indata Is Nothing Then
    Set wb_data = wb_indata
    Set ws_data = ws_indata
  ElseIf Not wb_outdata Is Nothing Then
    Set wb_data = wb_outdata
    Set ws_data = ws_outdata
  End If

  wb_data.Activate
  ws_data.Select

  max_col = ws_data.Cells(1, Columns.Count).End(xlToLeft).Column
  ReDim incol_array(max_col, 2)

  ' f[^t@CÌñÔðæ¾
  For i_cnt = 1 To max_col
    incol_array(i_cnt, 1) = ws_data.Cells(1, i_cnt)
    incol_array(i_cnt, 2) = i_cnt
  Next i_cnt

  wb.Activate
  ws_setup.Select

  max_row = ws_setup.Cells(Rows.Count, setup_col).End(xlUp).Row
  pass_flg = 0

  Range(Cells(setup_row, 3), Cells(max_row, 3)).ClearContents

  For setup_cnt = setup_row To max_row
    match_flg = 0
    If Mid(ws_setup.Cells(setup_cnt, setup_col), 1, 1) <> "*" Then
      For i_cnt = 1 To max_col
        If ws_setup.Cells(setup_cnt, setup_col) = incol_array(i_cnt, 1) Then
          ws_setup.Cells(setup_cnt, incol_col) = incol_array(i_cnt, 2)
          match_flg = 1
          Exit For
        End If
      Next i_cnt
      If pass_flg = 0 Then
        If match_flg = 0 Then
          MsgBox "QCODEm" & ws_setup.Cells(setup_cnt, 1) & "nªAf[^t@CÉ èÜ¹ñB", vbExclamation, "MCS 2020 - Datacol_Get"
          wb.Activate
          ws_mainmenu.Select
          Application.StatusBar = False
        End
      End If
    End If
  ElseIf ws_setup.Cells(setup_cnt, setup_col) = "*ÁHã" Then
    pass_flg = 1
    ws_setup.Cells(setup_cnt, incol_col) = incol_array(i_cnt, 2) + 1
  End If
Next setup_cnt

Application.ScreenUpdating = True
End Sub

Sub Setup_Hold()
  Dim max_row, max_col As Long

  '    Dim q_data() As question_data   ' QCODEÌf[^ðSÄæ¾
  Dim work_count As Long          ' ñJEgpÏ
  Dim work_string As String       ' ìÆp¶ñÏ
  Dim ctwork_count As Long        ' Ýâi[pJEgÏ
  '    Dim writing_target As Long      ' zñ«ÝÊui[pÏ

  '--------------------------------------------------------------------------------------------------'
  '@ÝèæÊÌîñðz[h@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@'
  '--------------------------------------------------------------------------------------------------'
  '@ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ì¬ú@2017.04.11@'
  '@ÅIÒWÒ@ºR@½@@@@@@@@@@@@@@@@@@@@@@@@@@@@ÒWú@2017.06.26@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "ÝèæÊîñ z[h..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_setup.Select
  '
  '    max_row = Cells(Rows.Count, setup_col).End(xlUp).Row
  '    max_col = Cells(1, Columns.Count).End(xlToLeft).Column
  '
  '    setup_array = Range(Cells(1, 1), Cells(max_row, max_col + 1))
  '

  ' --- 20170414 ºR½ ÇÁª ---------------------------------------------------------------------'

  ' QCODEÌðæ¾
  max_row = ws_setup.Cells(Rows.Count, setup_col).End(xlUp).Row

  ' QCODEÉí¹Äf[^zñðÄè`
  ReDim q_data(max_row)

  ' 20170419 QCODEÉí¹Äõpf[^zñðÄè`
  ReDim str_code(max_row)

  ' zñ«ÝÊuðú»
  'writing_target = 0

  ' QCODEÌîñðSÄæ¾·é
  For work_count = setup_row To max_row

    ' zñ«ÝÊuðCNg
    'writing_target = writing_target + 1

    ' 20170419 õpzñÉQCODEði[
    str_code(work_count) = ws_setup.Cells(work_count, 1).Value

    q_data(work_count).q_code = ws_setup.Cells(work_count, 1).Value                 ' QCODE
    q_data(work_count).r_code = ws_setup.Cells(work_count, 2).Value                 ' À
    q_data(work_count).m_code = ws_setup.Cells(work_count, 8).Value                 ' MCODE

    work_string = Trim(ws_setup.Cells(work_count, 9).Value)                         ' `®ðæ¾

    ' áÔA­ÔANA
    If Mid(work_string, 2, 1) = "M" Or _
      Mid(work_string, 2, 1) = "A" Or _
      Mid(work_string, 2, 1) = "C" Then

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT
      q_data(work_count).ct_loop = Val(Mid(work_string, 3))                   ' LCT
      q_data(work_count).q_format = Mid(work_string, 1, 2)                    ' `®

      ' LÌÝ
    Else

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT
      q_data(work_count).ct_loop = Val(Mid(work_string, 2))                   ' LCT
      q_data(work_count).q_format = Mid(work_string, 1, 1)                    ' `®

    End If

    ' `®É¶ÄðÏX
    Select Case q_data(work_count).q_format

      ' Tvio[âÇR[h
    Case "C"

      ' VOAT[
    Case "S"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT

      ' }`AT[
    Case "M"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT
      q_data(work_count).ct_loop = ws_setup.Cells(work_count, 16).Value       ' LCT


      ' AAT[
    Case "R"

      q_data(work_count).r_byte = Val(Mid(work_string, 2))                    ' ÀBYTE

      ' t[AT[
    Case "F"

      ' gJ[\
    Case "H"

      q_data(work_count).r_byte = 3                                           ' ÀBYTE

      ' I[vAT[
    Case "O"

      ' LMðÜßðsíÈ¢àÌ
    Case Else

  End Select

  q_data(work_count).r_unit = ws_setup.Cells(work_count, 4).Value                 ' ÀPÊi2017.05.10 ÇÁj

  q_data(work_count).sel_code1 = ws_setup.Cells(work_count, 10).Value             ' ZNg@ QCODE

  ' sel_code1ªüÍ³êÄ¢½
  If q_data(work_count).sel_code1 <> "" Then
    q_data(work_count).sel_value1 = Val(ws_setup.Cells(work_count, 11).Value)   ' ZNg@ VALUE
  End If

  q_data(work_count).sel_code2 = ws_setup.Cells(work_count, 12).Value             ' ZNgA QCODE

  ' sel_code2ªüÍ³êÄ¢½
  If q_data(work_count).sel_code2 <> "" Then
    q_data(work_count).sel_value2 = Val(ws_setup.Cells(work_count, 13).Value)   ' ZNgA VALUE
  End If

  q_data(work_count).sel_code3 = ws_setup.Cells(work_count, 14).Value             ' ZNgB QCODE

  ' sel_code3ªüÍ³êÄ¢½
  If q_data(work_count).sel_code3 <> "" Then
    q_data(work_count).sel_value3 = Val(ws_setup.Cells(work_count, 15).Value)   ' ZNgB VALUE
  End If

  ' OJeS[ðÜÞÝâ©ð»è
  '        If Trim(ws_setup.Cells(work_count, 17).Value) <> "" Then
  '            q_data(work_count).ct_0flg = True                                           ' 0CTtO
  '        End If

  q_data(work_count).q_title = ws_setup.Cells(work_count, 18).Value               ' \è

  ' Ýâª¶Ý·é
  If q_data(work_count).ct_count <> 0 Then

    ' JeS[ªÝâðæ¾·é
    For ctwork_count = 1 To q_data(work_count).ct_count

      q_data(work_count).q_ct(ctwork_count) = _
      ws_setup.Cells(work_count, 18 + ctwork_count).Value                     ' Ýâ

    Next ctwork_count

  End If

  ' üÍf[^ñÔªLü³êÄ¢é
  If ws_setup.Cells(work_count, 3).Value <> "" Then
    q_data(work_count).data_column = Val(ws_setup.Cells(work_count, 3).Value)   ' üÍf[^ñÔ
  End If

Next work_count

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub

'--------------------------------------------------------------------------------------------------'
' ì@¬@Ò ºR ½                                                           ì¬ú  2017.04.19  '
' ÅIÒWÒ ºR ½@@@@@@@@@@@@@@@@@@@@@@@@@@@   @ÒWú@2017.05.02@'
' QCODE}b`Opt@NV                                                                  '
' ø@ String^  QCODE                                                                           '
' ßèl Integer^ ñÔ                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_Match(ByVal match_code As String) As Integer

  ' G[ð³µÄðs¤
  On Error Resume Next

  ' QCODEðõµñÔðÔ·
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1

  ' QCODEª}b`OµÈ¢êÍG[\¦
  If Qcode_Match = 0 Then
    MsgBox "QCODEm" & match_code & "nÍÝèæÊÉ¶ÝµÜ¹ñB" & _
    vbCrLf & "vOð­§I¹µÜ·B", vbCritical, "MCS 2020 - Qcode_Matchy­§I¹z"
    wb.Activate
    ws_mainmenu.Select

    '2018/05/15 - ÇL ==========================
    Dim myWB As Workbook
    If Workbooks.Count > 1 Then
      Application.DisplayAlerts = False
      Application.Visible = True
      For Each myWB In Workbooks
        If myWB.Name <> ActiveWorkbook.Name Then
          myWB.Close
        End If
      Next
      Application.DisplayAlerts = True
    End If
    '============================================

    '2018/06/13 - ÇL ==========================
    ' SUMtH_àÌ0oCgt@Cðí
    Dim sum_file As String
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    Do Until sum_file = ""
      DoEvents
      If FileLen(file_path & "\SUM\" & sum_file) = 0 Then
        Kill file_path & "\SUM\" & sum_file
      End If
      sum_file = Dir()
    Loop
    '============================================
    Call Finishing_Mcs2017
  End
End If

End Function

'--------------------------------------------------------------------------------------------------'
' ì@¬@Ò ºR ½                                                           ì¬ú  2018.06.25  '
' ÅIÒWÒ ºR ½@@@@@@@@@@@@@@@@@@@@@@@@@@@   @ÒWú@2018.06.25@'
' QCODE}b`Opt@NViG[ßµo[Wj                                          '
' ø@ String^  QCODE                                                                           '
' ßèl Integer^ ñÔ                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_MatchE(ByVal match_code As String) As Integer
  ' G[ð³µÄðs¤
  On Error Resume Next
  ' QCODEðõµñÔðÔ·
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1
End Function

'--------------------------------------------------------------------------------------------------'
' ì¬Ò@@@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ì¬ú@2017.06.20@'
' ÅIÒWÒ@eè@m@@@@@@@@@@@@@@@@@@@@@@@@@@@@ ÒWú@2017.06.20@'
' Xe[^Xo[hbgANVpt@NV                                                   '
' ø@ Integer^ JE^[                                                                      '
' ßèl String^  hbg                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Status_Dot(ByVal i_cnt As Integer) As String
  If i_cnt Mod 4 = 0 Then
    Status_Dot = ""
  Else
    Status_Dot = String(i_cnt Mod 4, ".")
  End If
End Function

