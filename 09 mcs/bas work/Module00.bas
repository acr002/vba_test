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

' 20170419 村山誠 検索用QCODE配列
Public str_code() As String

' 20170425 村山誠 設定画面格納用ユーザ型
Public Type question_data
q_code As String                ' QCODE格納用文字列変数
r_code As String                ' 実数CODE格納用文字列変数
m_code As String                ' MCODE格納用文字列変数
q_format As String              ' 設問形式格納用文字列変数
r_byte As Integer               ' 実数桁数格納用変数
r_unit As String                ' 実数単位文字列変数（2017.05.10 追加）
sel_code1 As String             ' セレクト条件①QCODE格納用文字列変数
sel_value1 As Integer           ' セレクト条件①値格納用文字列変数
sel_code2 As String             ' セレクト条件②QCODE格納用文字列変数
sel_value2 As Integer           ' セレクト条件②値格納用文字列変数
sel_code3 As String             ' セレクト条件③QCODE格納用文字列変数
sel_value3 As Integer           ' セレクト条件③値格納用文字列変数
ct_count As Integer             ' 設問カテゴリー数格納用変数
ct_loop As Integer              ' ループカウント数格納用変数
'        ct_0flg As Boolean              ' ０カテゴリーフラグ格納用変数
q_title As String               ' 表題格納用文字列変数
q_ct(300) As String             ' 設問肢格納用文字列配列
data_column As Integer          ' 入力データ列番号格納用変数
End Type

Public q_data() As question_data    ' QCODE毎のデータを全て取得

public Sub Auto_Open()
  Dim base_pt As String
  '--------------------------------------------------------------------------------------------------'
  '　ファイルオープン処理　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.10　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03  '
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  '    Call Starting_Mcs2017    ' ThisWorkbook にて呼び出しに変更 - 2020.6.3
  Application.StatusBar = "ファイルオープン処理中..."
  Application.ScreenUpdating = False
  wb.Activate
  ws_mainmenu.Select
  If ws_mainmenu.Cells(gdrive_row, gdrive_col) <> "" And _
    ws_mainmenu.Cells(gcode_row, gcode_col) <> "" Then
    base_pt = ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If Dir(base_pt & "\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
      If (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 8) <> "// 読み込んだ日時") And _
        (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 7) <> "// 保存した日時") Then
        ws_mainmenu.Cells(initial_row, initial_col) = "// 初期設定済み"
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
'　シート構成の初期設定　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.10　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2017.04.14　'
'--------------------------------------------------------------------------------------------------'
public Sub Starting_Mcs2017()
  Set wb = ThisWorkbook
  Set ws_mainmenu = wb.Worksheets("メインメニュー")
  Set ws_setup = wb.Worksheets("設定画面")
  ' メインメニュー：業務コードの行列
  gcode_row = 3
  gcode_col = 8
  ' メインメニュー：作業ドライブの行列
  gdrive_row = 3
  gdrive_col = 23
  ' メインメニュー：初期設定済みメッセージ出力先行列
  initial_row = 6
  initial_col = 32
  ' 設定画面：先頭パラメータの行列
  setup_row = 3
  setup_col = 1
  ' 設定画面：入力データ列番号の行列
  incol_row = 3
  incol_col = 3
End Sub
'-----------------------------------------------------------------------------

public Sub Finishing_Mcs2017()
  '--------------------------------------------------------------------------------------------------'
  '　終了処理・各オブジェクトの参照を解除　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.10　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2018.09.19　'
  '--------------------------------------------------------------------------------------------------'
  Application.ScreenUpdating = True
  Application.StatusBar = False

  wb.Activate
  ws_setup.Select
  ws_setup.Cells(3, 1).Select
  ws_mainmenu.Select
  ws_mainmenu.Cells(3, 8).Select

  ' 設定画面チェックのエラーログファイルが0バイトならファイル削除
  If Dir(file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx") <> "" Then
    If FileLen(file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx") = 0 Then
      Kill file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx"
    End If
  End If

  ' ロジックチェックによるデータファイルのエラーログファイルが0バイトならファイル削除
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
  '　設定画面チェック　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.12　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.05.19　'
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  Call Starting_Mcs2017
  Application.StatusBar = "設定画面 チェック中..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If ws_mainmenu.Cells(gcode_row, gcode_col).Value = "" Then
    MsgBox "メインメニューの業務コードが未入力です。", vbExclamation, "MCS 2020 - Setup_Check"
    Cells(gcode_row, gcode_col).Select
    Application.StatusBar = False
    End
  End If

  Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & _
  "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & "_mcs.ini" For Input As #1
  Line Input #1, file_path
  Line Input #1, setup_status
  Close #1

  Open file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx" For Append As #1
  Close #1
  If Err.Number > 0 Then
    Workbooks(ope_code & "_設定画面NG.xlsx").Close
  End If

  If Dir(file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx") <> "" Then
    Kill file_path & "\4_LOG\" & ope_code & "_設定画面NG.xlsx"
  End If

  ws_setup.Select
  max_row = Cells(Rows.Count, setup_col).End(xlUp).Row

  ' ここからエラーチェックのコーディング(´･ω･`)
  err_msg = "【設定画面のエラーリスト】" & vbCrLf & "QCODE,行数,エラー項目,エラー内容" & vbCrLf

  '★サンプルナンバーのQCODEチェック
  If Cells(3, 1).Value <> "SNO" Then
    err_msg = err_msg & Cells(3, 1).Value & ",3行目" & ",QCODE(RC1)" & ",サンプル№のQCODEは［SNO］としてください。" & vbCrLf
  End If

  For qcode_cnt = setup_row To max_row
    If Cells(qcode_cnt, setup_col).Value = "*加工後" Then
      Rows(qcode_cnt).Select
      With Selection.Interior
        .Color = 65535
      End With
    ElseIf Left(Cells(qcode_cnt, setup_col).Value, 1) = "*" Then
      '処理無し
    Else
      '設定画面行列情報登録
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

    '★QCODE の重複チェック
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
    Cells(qcode_cnt, setup_col).Value)
    If err_cnt >= 2 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",QCODE(RC1)" & ",QCODEは重複不可です。" & vbCrLf
    End If

    '★QCODE未指定かつ非コメント行への入力チェック
    If Cells(qcode_cnt, setup_col).Value = "" And _
      WorksheetFunction.CountA(Range(Cells(qcode_cnt, setup_col), Cells(qcode_cnt, max_col))) <> 0 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",QCODE(RC1)" & ",QCODEがブランクの場合、他の設定項目への入力は不可となります。" & vbCrLf
    End If

    '★MCODE単独入力チェック
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, mcode_col), Cells(max_row, mcode_col)), _
    Cells(qcode_cnt, mcode_col).Value)
    If err_cnt = 1 And Cells(qcode_cnt, mcode_col).Value <> "" Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",MCODE(RC8)" & ",MCODEは2つ以上の設問に対して指定してください。" & vbCrLf
    End If

    '★形式チェック
    qformat_typ = Cells(qcode_cnt, qCStr_col).Value
    Select Case Left(qformat_typ, 1)
      Case "C", "S", "M", "H", "F", "O"
        If Len(qformat_typ) > 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",形式は1桁で［C］［S］［M］［H］［F］［O］を指定してください。" & vbCrLf
        End If
      Case "R"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",形式［R］は桁数を指定してください。" & vbCrLf
        ElseIf Len(qformat_typ) > 3 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",桁数は2桁まで指定可能です。" & vbCrLf
        ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",桁数は半角数字で指定してください。" & vbCrLf
        End If
      Case "L"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",形式 ［L］は回答制限数の指定をしてください。" & vbCrLf
        ElseIf (Mid(qformat_typ, 2, 1) = "M") Or (Mid(qformat_typ, 2, 1) = "A") Or (Mid(qformat_typ, 2, 1) = "C") Then
          If Len(qformat_typ) = 2 Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",形式 ［L］は回答制限数の指定をしてください。" & vbCrLf
          ElseIf Len(qformat_typ) > 5 And IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",回答制限数は3桁まで指定可能です。" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",回答制限数は半角数字で指定してください。" & vbCrLf
          End If
        ElseIf Val(Mid(qformat_typ, 2, 1)) = 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",リミットマルチの形式は［LMx］［LAx］［LCx］のいずれかで指定をしてください（xは回答数）。" & vbCrLf
        Else
          If Len(qformat_typ) > 4 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",回答制限数は3桁まで指定可能です。" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",回答制限数は半角数字で指定してください。" & vbCrLf
          End If
        End If
      Case Else
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",形式(RC9)" & ",形式［C］［S］［M］［H］［F］［O］［R］［L］以外の指定は不可となります。" & vbCrLf
    End Select

    '★セレクトチェック
    err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3)))
    If err_cnt <> 0 Then
      '★セレクト条件記法チェック
      If WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_1))) < 2 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_2), Cells(qcode_cnt, selval_col_2))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①②③(RC10:RC15)" & ",条件（セレクト）の指定はQCODEと値のセットを条件①から左詰めで指定してください。" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_2))) < 4 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①②③(RC10:RC15)" & ",条件（セレクト）の指定はQCODEと値のセットを条件①から左詰めで指定してください。" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3))) < 6 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①②③(RC10:RC15)" & ",条件（セレクト）の指定はQCODEと値のセットを条件①から左詰めで指定してください。" & vbCrLf
      End If

      '★セレクト条件QCODE該当チェック
      '条件１
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_1).Value) = 0 And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件（セレクト）となるQCODEは予めA列に指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        find_row_1 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_1).Value, lookat:=xlWhole).Row
      End If
      '条件2
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_2).Value) = 0 And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件（セレクト）となるQCODEは予めA列に指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        find_row_2 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_2).Value, lookat:=xlWhole).Row
      End If
      '条件3
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_3).Value) = 0 And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件（セレクト）となるQCODEは予めA列に指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        find_row_3 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_3).Value, lookat:=xlWhole).Row
      End If

      '★セレクト条件値チェック
      '条件１
      If Cells(qcode_cnt, selval_col_1).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件（セレクト）の値は半角数字で指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value = "" And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value <> "" And Cells(qcode_cnt, selcode_col_1).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = True And find_row_1 <> 0 Then
        If Cells(find_row_1, ct_cnt_col).Value < Val(Cells(qcode_cnt, selval_col_1).Value) Then
          Debug.Print find_row_1
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件（セレクト）の値は該当設問の選択肢数以下の数値で指定してください。" & vbCrLf
        End If
      End If
      '条件２
      If Cells(qcode_cnt, selval_col_2).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件（セレクト）の値は半角数字で指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value = "" And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value <> "" And Cells(qcode_cnt, selcode_col_2).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = True And find_row_2 <> 0 Then
        If Cells(find_row_2, ct_cnt_col).Value - Cells(find_row_2, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_2).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件（セレクト）の値は該当設問の選択肢数以下の数値で指定してください。" & vbCrLf
        End If
      End If
      '条件３
      If Cells(qcode_cnt, selval_col_3).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件（セレクト）の値は半角数字で指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value = "" And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value <> "" And Cells(qcode_cnt, selcode_col_3).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件（セレクト）はQCODEと値のセットで指定してください。" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = True And find_row_3 <> 0 Then
        If Cells(find_row_3, ct_cnt_col).Value - Cells(find_row_3, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_3).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件（セレクト）の値は該当設問の選択肢数以下の数値で指定してください。" & vbCrLf
        End If
      End If
    End If

    '★選択肢数チェック
    Select Case Left(Cells(qcode_cnt, 9).Value, 1)
      Case "S", "M", "L"
        If max_col >= ct_st_col Then
          err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, ct_st_col), Cells(qcode_cnt, max_col)))
          If err_cnt <> Cells(qcode_cnt, ct_cnt_col).Value Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",CT数・カテゴリー内容(RC16＆RC19～)" & ",P列のCT数と、S列以降のカテゴリー項目の個数が一致しません。" & vbCrLf
          ElseIf err_cnt <> max_col - qttl_col Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",カテゴリー項目(RC19～)" & ",S列以降のカテゴリー項目は左詰めで入力してください。" & vbCrLf
          End If
        End If
      Case "F", "O", "R", "H", "C"
        If max_col >= ct_st_col Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",カテゴリー項目(RC19～)" & ",形式［R］［H］［F］［O］［C］のカテゴリー項目の設定はできません。" & vbCrLf
        ElseIf Cells(qcode_cnt, ct_cnt_col).Value <> 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",CT数(RC16)" & ",形式［R］［H］［F］［O］［C］の設問のCT数は［ブランク］にしてください。" & vbCrLf
        End If
      Case Else
    End Select

    ' 2020.1.9 - ゼロフラグはなくなりましたので、コメントアウトしました。
    '            '★ゼロフラグチェック
    '            If Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Len(Cells(qcode_cnt, ct_cnt_col).Value) = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",0CT(RC17)" & ",0CTのフラグはCT数が［1以上］の設問にのみ設定可能です。" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, ct_cnt_col).Value = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",0CT(RC17)" & ",0CTのフラグはCT数が［1以上］の設問にのみ設定可能です。" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, zero_f_col).Value <> 1 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",0CT(RC17)" & ",0CTの指定として使えるのは［1］のみとなります。" & vbCrLf
    '            End If

    '★設問タイトル有無チェック
    If Cells(qcode_cnt, qttl_col) = "" Then
      If Cells(qcode_cnt, ct_st_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",表題(RC18)" & ",設問の表題は必ず指定してください。" & vbCrLf
      End If
    End If

    '★予備エリアへの入力チェック
    For sys_i = 1 To sys_cnt
      sys_col = sys_num + sys_i - 1
      If Cells(qcode_cnt, sys_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",System(RC5:RC7)" & ",現在のバージョンではE~G列は使用しないでください。" & vbCrLf
      End If
    Next sys_i

    '★実数指定設問と該当実数設問のセレクト状況の同一チェック
    If Cells(qcode_cnt, rcode_col).Value <> "" Then
      Set rcode_cell = ws_setup.Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
      Find(What:=Cells(qcode_cnt, rcode_col).Value)
      If rcode_cell Is Nothing Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",実数(RC2)" & ",実数指定されたQCODEが見つかりません。" & vbCrLf
      Else
        rcode_row = rcode_cell.Row
        If Cells(qcode_cnt, selcode_col_1).Value <> Cells(rcode_row, selcode_col_1).Value Or _
          Cells(qcode_cnt, selval_col_1).Value <> Cells(rcode_row, selval_col_1).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件①(RC10:RC11)" & ",条件①（セレクト）の指定が該当の実数設問のQCODE［" _
          & Cells(rcode_row, setup_col).Value & "］と一致しません。" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_2).Value <> Cells(rcode_row, selcode_col_2).Value Or _
          Cells(qcode_cnt, selval_col_2).Value <> Cells(rcode_row, selval_col_2).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件②(RC12:RC13)" & ",条件②（セレクト）の指定が該当の実数設問のQCODE［" _
          & Cells(rcode_row, setup_col).Value & "］と一致しません。" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_3).Value <> Cells(rcode_row, selcode_col_3).Value Or _
          Cells(qcode_cnt, selval_col_3).Value <> Cells(rcode_row, selval_col_3).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "行目" & ",条件③(RC14:RC15)" & ",条件③（セレクト）の指定が該当の実数設問のQCODE［" _
          & Cells(rcode_row, setup_col).Value & "］と一致しません。" & vbCrLf
        End If
      End If
    End If
  End If
Next qcode_cnt

Cells(setup_row, setup_col).Select

Application.ScreenUpdating = True
preset_gcode = ws_mainmenu.Cells(gcode_row, gcode_col).Value
If err_msg <> "【設定画面のエラーリスト】" & vbCrLf & "QCODE,行数,エラー項目,エラー内容" & vbCrLf Then
  Application.DisplayAlerts = False
  ' エラーメッセージ出力ファイルの作成
  Workbooks.Add
  Set wb_preset = ActiveWorkbook
  Set ws_preset = wb_preset.Worksheets("Sheet1")
  Columns("A").NumberFormat = "@"
  If Dir(file_path & "\4_LOG\" & preset_gcode & "_設定画面NG.xlsx") = "" Then
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_設定画面NG.xlsx"
  Else
    Kill file_path & "\4_LOG\" & preset_gcode & "_設定画面NG.xlsx"
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_設定画面NG.xlsx"
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
    .Name = "Takaoゴシック"
    .Size = 11
  End With
  Columns("B:D").EntireColumn.AutoFit
  ws_preset.Cells(1, 1).Select
  wb_preset.Save
  MsgBox "設定画面の内容にエラーがあります。" & vbCrLf & _
  file_path & "\4_LOG\" & preset_gcode & "_設定画面NG.xlsx を確認してください。", vbExclamation, "MCS 2020 - Setup_Check"
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
  '　ファイルパスの取得　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.10　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2017.05.18　'
  '--------------------------------------------------------------------------------------------------'
  Application.StatusBar = "ファイルパス取得・生成中..."
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
    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］が見つかりません。" _
    & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Filepath_Get"
    End
  End If
  Application.ScreenUpdating = True
  Application.StatusBar = False
End Sub

Sub Indata_Open()
  Dim indata_fn, revdata_fn As String
  '--------------------------------------------------------------------------------------------------'
  '　入力データのオープン        　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.11　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.04.08　'
  '--------------------------------------------------------------------------------------------------'
  '【概要】　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '　このプロシージャでオープンされるファイルは、以下のパターンのファイルのみ　　　　　　　　　　　　'
  '　・入力データファイル ...... *IN.xlsx　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '　・修正後データファイル .... *RE.xlsx                  　　　　　　　　　　　　　　　　　　　　　'
  '　・Call元：入力データの修正［Module04］                                                          '
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "入力データ オープン中..."
  '    Application.ScreenUpdating = False    ' ファイルオープンの進捗が見られた方がよいのでコメントアウト

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
    MsgBox "入力データファイルが存在しません。", vbExclamation, "MCS 2020 - Indata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Outdata_Open()
  '--------------------------------------------------------------------------------------------------'
  '　加工後データのオープン　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.14　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.04.08　'
  '--------------------------------------------------------------------------------------------------'
  '【概要】　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '　このプロシージャでオープンされるファイルは、集計設定ファイルで指定されているファイル　　　　　　'
  '　Call元：集計サマリーデータの作成［Module52］　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "加工後データ オープン中..."
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
    MsgBox "加工後データファイル［" & outdata_fn & "］が存在しません。", vbExclamation, "MCS 2020 - Outdata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Datafile_Open()
  '--------------------------------------------------------------------------------------------------'
  '　データファイルのオープン　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.27　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.04.08　'
  '--------------------------------------------------------------------------------------------------'
  '【概要】　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '　このプロシージャでオープンされるファイルは、ユーザーから入力されたファイル　　　　　　　　　　　'
  '　Call元：入力データのロジックチェック［Module05］　　　　　　　　　　　　　　　　　　　　　　　　'
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
    MsgBox "指定したファイル［" & outdata_fn & "］が存在しません。", vbExclamation, "MCS 2020 - Datafile_Open"
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
  '　データファイルの列番号を取得　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.27　'
  '　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2018.04.04　'
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

  ' データファイルの列番号を取得
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
          MsgBox "QCODE［" & ws_setup.Cells(setup_cnt, 1) & "］が、データファイルにありません。", vbExclamation, "MCS 2020 - Datacol_Get"
          wb.Activate
          ws_mainmenu.Select
          Application.StatusBar = False
        End
      End If
    End If
  ElseIf ws_setup.Cells(setup_cnt, setup_col) = "*加工後" Then
    pass_flg = 1
    ws_setup.Cells(setup_cnt, incol_col) = incol_array(i_cnt, 2) + 1
  End If
Next setup_cnt

Application.ScreenUpdating = True
End Sub

Sub Setup_Hold()
  Dim max_row, max_col As Long

  '    Dim q_data() As question_data   ' QCODE毎のデータを全て取得
  Dim work_count As Long          ' 処理回数カウント用変数
  Dim work_string As String       ' 作業用文字列変数
  Dim ctwork_count As Long        ' 設問肢格納用カウント変数
  '    Dim writing_target As Long      ' 配列書き込み位置格納用変数

  '--------------------------------------------------------------------------------------------------'
  '　設定画面の情報をホールド　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
  '--------------------------------------------------------------------------------------------------'
  '　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.11　'
  '　最終編集者　村山　誠　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2017.06.26　'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "設定画面情報 ホールド中..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_setup.Select
  '
  '    max_row = Cells(Rows.Count, setup_col).End(xlUp).Row
  '    max_col = Cells(1, Columns.Count).End(xlToLeft).Column
  '
  '    setup_array = Range(Cells(1, 1), Cells(max_row, max_col + 1))
  '

  ' --- 20170414 村山誠 追加部分 ---------------------------------------------------------------------'

  ' QCODEの数を取得
  max_row = ws_setup.Cells(Rows.Count, setup_col).End(xlUp).Row

  ' QCODE数に合わせてデータ配列数を再定義
  ReDim q_data(max_row)

  ' 20170419 QCODE数に合わせて検索用データ配列数を再定義
  ReDim str_code(max_row)

  ' 配列書き込み位置を初期化
  'writing_target = 0

  ' QCODEの情報を全て取得する
  For work_count = setup_row To max_row

    ' 配列書き込み位置をインクリメント
    'writing_target = writing_target + 1

    ' 20170419 検索用配列にQCODEを格納
    str_code(work_count) = ws_setup.Cells(work_count, 1).Value

    q_data(work_count).q_code = ws_setup.Cells(work_count, 1).Value                 ' QCODE
    q_data(work_count).r_code = ws_setup.Cells(work_count, 2).Value                 ' 実数
    q_data(work_count).m_code = ws_setup.Cells(work_count, 8).Value                 ' MCODE

    work_string = Trim(ws_setup.Cells(work_count, 9).Value)                         ' 形式を取得

    ' 若番、強番、クリア
    If Mid(work_string, 2, 1) = "M" Or _
      Mid(work_string, 2, 1) = "A" Or _
      Mid(work_string, 2, 1) = "C" Then

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT数
      q_data(work_count).ct_loop = Val(Mid(work_string, 3))                   ' LCT数
      q_data(work_count).q_format = Mid(work_string, 1, 2)                    ' 形式

      ' Lのみ
    Else

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT数
      q_data(work_count).ct_loop = Val(Mid(work_string, 2))                   ' LCT数
      q_data(work_count).q_format = Mid(work_string, 1, 1)                    ' 形式

    End If

    ' 形式に応じて処理を変更
    Select Case q_data(work_count).q_format

      ' サンプルナンバーや管理コード等
    Case "C"

      ' シングルアンサー
    Case "S"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT数

      ' マルチアンサー
    Case "M"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT数
      q_data(work_count).ct_loop = ws_setup.Cells(work_count, 16).Value       ' LCT数


      ' リアルアンサー
    Case "R"

      q_data(work_count).r_byte = Val(Mid(work_string, 2))                    ' 実数BYTE数

      ' フリーアンサー
    Case "F"

      ' Ｈカーソル
    Case "H"

      q_data(work_count).r_byte = 3                                           ' 実数BYTE数

      ' オープンアンサー
    Case "O"

      ' LMを含め処理を行わないもの
    Case Else

  End Select

  q_data(work_count).r_unit = ws_setup.Cells(work_count, 4).Value                 ' 実数単位（2017.05.10 追加）

  q_data(work_count).sel_code1 = ws_setup.Cells(work_count, 10).Value             ' セレクト① QCODE

  ' sel_code1が入力されていた時
  If q_data(work_count).sel_code1 <> "" Then
    q_data(work_count).sel_value1 = Val(ws_setup.Cells(work_count, 11).Value)   ' セレクト① VALUE
  End If

  q_data(work_count).sel_code2 = ws_setup.Cells(work_count, 12).Value             ' セレクト② QCODE

  ' sel_code2が入力されていた時
  If q_data(work_count).sel_code2 <> "" Then
    q_data(work_count).sel_value2 = Val(ws_setup.Cells(work_count, 13).Value)   ' セレクト② VALUE
  End If

  q_data(work_count).sel_code3 = ws_setup.Cells(work_count, 14).Value             ' セレクト③ QCODE

  ' sel_code3が入力されていた時
  If q_data(work_count).sel_code3 <> "" Then
    q_data(work_count).sel_value3 = Val(ws_setup.Cells(work_count, 15).Value)   ' セレクト③ VALUE
  End If

  ' ０カテゴリーを含む設問かを判定
  '        If Trim(ws_setup.Cells(work_count, 17).Value) <> "" Then
  '            q_data(work_count).ct_0flg = True                                           ' 0CTフラグ
  '        End If

  q_data(work_count).q_title = ws_setup.Cells(work_count, 18).Value               ' 表題

  ' 設問肢が存在する時
  If q_data(work_count).ct_count <> 0 Then

    ' カテゴリー数分設問肢を取得する
    For ctwork_count = 1 To q_data(work_count).ct_count

      q_data(work_count).q_ct(ctwork_count) = _
      ws_setup.Cells(work_count, 18 + ctwork_count).Value                     ' 設問肢

    Next ctwork_count

  End If

  ' 入力データ列番号が記入されている時
  If ws_setup.Cells(work_count, 3).Value <> "" Then
    q_data(work_count).data_column = Val(ws_setup.Cells(work_count, 3).Value)   ' 入力データ列番号
  End If

Next work_count

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.04.19  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.02　'
' QCODEマッチング用ファンクション                                                                  '
' 引　数 String型  QCODE                                                                           '
' 戻り値 Integer型 列番号                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_Match(ByVal match_code As String) As Integer

  ' エラーを無視して処理を行う
  On Error Resume Next

  ' QCODEを検索し列番号を返す
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1

  ' QCODEがマッチングしない場合はエラー表示
  If Qcode_Match = 0 Then
    MsgBox "QCODE［" & match_code & "］は設定画面に存在しません。" & _
    vbCrLf & "プログラムを強制終了します。", vbCritical, "MCS 2020 - Qcode_Match【強制終了】"
    wb.Activate
    ws_mainmenu.Select

    '2018/05/15 - 追記 ==========================
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

    '2018/06/13 - 追記 ==========================
    ' SUMフォルダ内の0バイトファイルを削除
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
' 作　成　者 村山 誠                                                           作成日  2018.06.25  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2018.06.25　'
' QCODEマッチング用ファンクション（エラー戻しバージョン）                                          '
' 引　数 String型  QCODE                                                                           '
' 戻り値 Integer型 列番号                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_MatchE(ByVal match_code As String) As Integer
  ' エラーを無視して処理を行う
  On Error Resume Next
  ' QCODEを検索し列番号を返す
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1
End Function

'--------------------------------------------------------------------------------------------------'
' 作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　 作成日　2017.06.20　'
' 最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　 編集日　2017.06.20　'
' ステータスバードットアクション用ファンクション                                                   '
' 引　数 Integer型 カウンター                                                                      '
' 戻り値 String型  ドット                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Status_Dot(ByVal i_cnt As Integer) As String
  If i_cnt Mod 4 = 0 Then
    Status_Dot = ""
  Else
    Status_Dot = String(i_cnt Mod 4, ".")
  End If
End Function

