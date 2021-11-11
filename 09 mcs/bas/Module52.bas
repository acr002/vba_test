Attribute VB_Name = "Module52"
Option Explicit
    Public wb_tabinst As Workbook
    Public ws_tabinst As Worksheet
    
    Public wb_summary As Workbook
    Public ws_summary0 As Worksheet
    Public ws_summary1 As Worksheet
    Public ws_summary2 As Worksheet
    Public ws_summary3 As Worksheet
    
    Public tabinst_fn As String
    Public summary_fn As String
    Public summary_fd As String
    
    Dim c_cnt As Long
    Dim outdata_row As Long             ' 集計データの最終行
    Dim outdata_col As Long             ' 集計データの最終列
    
    Dim a_index, f_index As Long        ' 表頭・表側の設定情報のインデックス
    Dim ra_index As Long                ' 実数設問のインデックス
    
    ' 表頭・表側・集計設定に、セレクトの設定があるかないかの判定用です。
    ' データがセレクト条件を満たしているかは、16384列目で判定します。
    Dim select_flg As Integer           ' セレクト設定の有無フラグ
    
    ' 2019.10.10 - ウエイト集計関連追加
    Dim weight_flg As String            ' ウエイト集計有無のフラグ（なし、あり）
    Dim weight_col As Long              ' 補正値の列
    Dim w_index As Long                 ' 補正値のインデックス
    
    Dim se_index As Long                ' 集計設定セレクトのインデックス
    Dim as1_index As Long               ' 表頭セレクトのインデックス
    Dim as2_index As Long
    Dim as3_index As Long
    Dim fs1_index As Long               ' 表側セレクトのインデックス
    Dim fs2_index As Long
    Dim fs3_index As Long
    
    Dim ama_cnt As Long                 ' 表頭のＭＡカテゴリー数（0:SA、0以外:MA）
    Dim fma_cnt As Long                 ' 表側のＭＡカテゴリー数（0:SA、0以外:MA）
    
    Dim se_cnt As Long                  ' 集計設定セレクトのＭＡカテゴリー数（0:SA、0以外:MA）
    Dim as1_cnt As Long                 ' 表頭セレクトのＭＡカテゴリー数（0:SA、0以外:MA）
    Dim as2_cnt As Long
    Dim as3_cnt As Long
    Dim fs1_cnt As Long                 ' 表側セレクトのＭＡカテゴリー数（0:SA、0以外:MA）
    Dim fs2_cnt As Long
    Dim fs3_cnt As Long
    
    Dim se_msg As String                ' 集計設定セレクトのメッセージ
    Dim as1_msg As String               ' 表頭セレクトのメッセージ
    Dim as2_msg As String
    Dim as3_msg As String
    Dim fs1_msg As String               ' 表側セレクトのメッセージ
    Dim fs2_msg As String
    Dim fs3_msg As String
    
    Dim sum_row As Long                 ' 集計表（Ｎ％表）の行列
    Dim sum_col As Long                 ' 集計表（Ｎ％表）の行列
    
    Dim div_row As Long                 ' 集計表（Ｎ表、％表）の行列
    Dim div_col As Long                 ' 集計表（Ｎ表、％表）の行列
    
    Public Type cross_data
        hyo_num As String               ' 表№格納用文字列変数
        f_code As String                ' 表側QCODE格納用文字列変数
        a_code As String                ' 表頭QCODE格納用文字列変数
        r_code As String                ' 実数設問QCODE格納用文字列変数
        fna_flg As String               ' 表側表示フラグ格納用文字列変数
        ana_flg As String               ' 表頭NAフラグ格納用文字列変数
        bosu_flg As String              ' 集計母数フラグ格納用変数
    
        sum_flg As String               ' 実数設問出力指示・合計フラグ格納用文字列変数
        ave_flg As String               ' 実数設問出力指示・平均フラグ格納用文字列変数
        sd_flg As String                ' 実数設問出力指示・標準偏差フラグ格納用文字列変数
        min_flg As String               ' 実数設問出力指示・最小値フラグ格納用文字列変数
        q1_flg As String                ' 実数設問出力指示・第１四分位フラグ格納用文字列変数
        med_flg As String               ' 実数設問出力指示・中央値フラグ格納用文字列変数
        q3_flg As String                ' 実数設問出力指示・第３四分位フラグ格納用文字列変数
        max_flg As String               ' 実数設問出力指示・最大値フラグ格納用文字列変数
        mod_flg As String               ' 実数設問出力指示・最頻値フラグ格納用文字列変数
    
        sel_code As String              ' セレクト条件・QCODE
        sel_value As Integer            ' セレクト条件・値
        
        ken_flg As String               ' 表示オプション・件数欄フラグ
        yuko_flg As String              ' 表示オプション・有効回答フラグ
        nobe_flg As String              ' 表示オプション・延べ回答
    
        top1_flg As String              ' TOP1・マーキングフラグ
        sort_flg As String              ' CTソート・降順フラグ
        exct_flg As String              ' CTソート・除外CTフラグ
        graph_flg As String             ' グラフ・作成フラグ
    End Type
    
    Dim c_data() As cross_data          ' 表№毎の集計指示を全て取得

Sub Summarydata_Creation()
    Dim yensign_pos As Long
    Dim r_code As Integer
'2018/05/23 - 追記 ==========================
    Dim crs_tab() As String
    Dim crs_file As String
    Dim crs_cnt As Long
    Dim i_cnt As Long
    Dim fn_cnt As Long
'--------------------------------------------------------------------------------------------------'
'　集計サマリーデータの作成　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.05.22　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
'【戯れ言】2017.05.10　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'　サマリーファイル作成前にロジックチェックを組み込むか検討しましたが、集計設定ファイルが複数ある　'
'　場合などのケースで、チェック１回、集計が複数回のときに毎回同じチェックをすると時間がかかるので、'
'　モジュールは独立したものとします。　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "集計サマリーデータの作成 初期化作業中..."
    
    Open file_path & "\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面NG.xlsx" For Append As #1
    Close #1
    If Err.Number > 0 Then
        MsgBox "設定画面エラーファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面NG.xlsx］が開かれています。" _
        & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "設定画面の入力情報にエラーがある可能性があります。" _
        & vbCrLf & "エラーの内容を確認して［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面NG.xlsx］を閉じてから、" _
        & vbCrLf & "再実行してください。" _
        , vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    Open file_path & "\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx" For Append As #1
    Close #1
    If Err.Number > 0 Then
        MsgBox "ロジックチェックのエラーログファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx］が開かれています。" _
        & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "ロジックチェックを行ったデータにエラーがある可能性があります。" _
        & vbCrLf & "エラーの内容を確認して［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx］を閉じてから、" _
        & vbCrLf & "再実行してください。" _
        , vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ' 事前チェックのメッセージ
    wb.Activate
    ws_mainmenu.Select
    r_code = MsgBox("事前に集計するファイルのロジックチェックを行いましたか。" _
    & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "事前にロジックチェックを行う対象ファイルは、" _
    & vbCrLf & "集計設定ファイルの［集計するファイル名］の項目で" & vbCrLf & "指定したファイルになります。", _
    vbYesNo + vbQuestion, "MCS 2020 - Summarydata_Creation")
    If r_code = vbNo Then
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ChDrive file_path
    ChDir file_path & "\3_FD"
    
    ' CRSフォルダ内のxlsx形式のファイル数をカウント
    crs_cnt = 0
    crs_file = Dir(file_path & "\3_FD\CRS\*.xlsx")
    Do Until crs_file = ""
        DoEvents
        crs_cnt = crs_cnt + 1
        crs_file = Dir()
    Loop
    
    ' CRSフォルダ内のxlsx形式のファイル名を配列にセット
    ReDim crs_tab(crs_cnt)
    crs_file = Dir(file_path & "\3_FD\CRS\*.xlsx")
    For fn_cnt = 1 To crs_cnt
        DoEvents
        crs_tab(fn_cnt) = crs_file
        crs_file = Dir()
    Next fn_cnt
    fn_cnt = crs_cnt
    
' 集計サマリーデータ複数回作成処理
    If crs_cnt > 0 Then
        r_code = MsgBox("FDフォルダ内に［CRS］フォルダがあります。" _
         & vbCrLf & "CRSフォルダ内にある" & fn_cnt & "個の集計設定ファイルを使用して、" _
         & vbCrLf & "集計サマリーデータを一括作成しますか。" _
         & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "CRSフォルダ内の［xlsx形式］のファイル数を表示しています。" _
         & vbCrLf & "「はい」　→ 集計サマリーデータを一括作成" & vbCrLf & "「いいえ」→ 集計設定ファイルを選択してから作成", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Summarydata_Creation")
        If r_code = vbYes Then
            crs_cnt = 1
            For i_cnt = 1 To fn_cnt
                DoEvents
                tabinst_fn = crs_tab(i_cnt)
                wb.Activate
                ws_mainmenu.Select
                If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\CRS\" & tabinst_fn) <> "" Then
                    Workbooks.Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & _
                     ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\CRS\" & tabinst_fn
                    Set wb_tabinst = Workbooks(tabinst_fn)
                    Set ws_tabinst = wb_tabinst.ActiveSheet
                    If ws_tabinst.Cells(2, 4) = "" Then
                        MsgBox "集計設定ファイル［" & tabinst_fn & "］で、集計するデータファイルが入力されていません。", vbExclamation, "MCS 2020 - Summarydata_Creation"
                        Application.StatusBar = False
                        End
                    End If
                    outdata_fn = ws_tabinst.Cells(2, 4)
                    weight_flg = ws_tabinst.Cells(2, 12)    ' ウエイト集計の有無フラグ取得
                    If weight_flg = "" Then weight_flg = "なし"
                Else
                    MsgBox "集計設定ファイル［" & tabinst_fn & "］が存在しません。", vbExclamation, "MCS 2020 - Summarydata_Creation"
                    Application.StatusBar = False
                    End
                End If
                
                Call Outdata_Open
                Call Setup_Hold
                
                wb_outdata.Activate
                ws_outdata.Select
                outdata_row = Cells(Rows.Count, setup_col).End(xlUp).Row
                outdata_col = Cells(1, Columns.Count).End(xlToLeft).Column
                
                If weight_flg = "あり" Then   ' ウエイト（補正値）の取得
                  w_index = Qcode_Match("weight")
                End If
                
                ' 2018/05/28 - サマリーの出力先は［SUM］フォルダ固定に変更。
'                summary_fd = Replace(Left(tabinst_fn, InStr(tabinst_fn, ".") - 1), ws_mainmenu.Cells(gcode_row, gcode_col), "")
                summary_fd = "SUM"
                summary_fn = Left(tabinst_fn, InStr(tabinst_fn, ".") - 1) & "_sum.xlsx"
                
'2018/04/26 - 追記 ==========================
                Open file_path & "\" & summary_fd & "\" & summary_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(summary_fn).Close
                End If
                If Dir(file_path & "\" & summary_fd & "\" & summary_fn) <> "" Then
                    Kill file_path & "\" & summary_fd & "\" & summary_fn
                End If
'============================================
                
                Call Cross_Setting(crs_cnt, fn_cnt)
                
                ' データファイルのクローズ
                wb_outdata.Activate
                ws_outdata.Select
                Columns("XFA:XFD").Select
                Selection.ClearContents
                Range("B7").Select
                Application.DisplayAlerts = False
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                Set wb_outdata = Nothing
                Set ws_outdata = Nothing
                
                ' 集計設定ファイルのクローズ
                wb_tabinst.Activate
                Application.DisplayAlerts = False
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                Set wb_tabinst = Nothing
                Set ws_tabinst = Nothing
                
'2020/01/10 - MCODE処理追記 =================
                Call Mcode_Setting
'============================================
    
                ' 集計サマリーファイルを保存してクローズ
                wb_summary.Activate
                ws_summary1.Select
                
                If Dir(file_path & "\" & summary_fd, vbDirectory) = "" Then
                    MkDir file_path & "\" & summary_fd
                End If
                
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=file_path & "\" & summary_fd & "\" & summary_fn
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                
                crs_cnt = crs_cnt + 1
            Next i_cnt
            
            Set wb_summary = Nothing
            Set ws_summary0 = Nothing
            Set ws_summary1 = Nothing
            Set ws_summary2 = Nothing
            Set ws_summary3 = Nothing
            
            wb.Activate
            ws_setup.Select
            ws_setup.Cells(1, 1).Select
            ws_mainmenu.Select
            ws_mainmenu.Cells(3, 8).Select
            
' システムログの出力
            ' 2020.6.3 - 追加
            ActiveSheet.Unprotect Password:=""
            ws_mainmenu.Cells(initial_row, initial_col).Locked = False
            If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
              ws_mainmenu.Cells(41, 6) = "23"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 23"
            End If
            ws_mainmenu.Cells(initial_row, initial_col).Locked = True
            ActiveSheet.Protect Password:=""
    
            Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
             "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
            Close #1
            If Err.Number > 0 Then
                Close #1
            End If
            If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
             "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his") = "" Then
                Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
                 "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Output As #1
                Print #1, ws_mainmenu.Cells(gcode_row, gcode_col) & " MCS 2020 operation history"
                Close #1
            End If
            Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
             "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計サマリーデータの作成：対象ファイル［SUMフォルダ内の" & crs_cnt - 1 & "個の集計設定ファイル］"
            Close #1
            Call Finishing_Mcs2017
            MsgBox crs_cnt - 1 & "個の集計サマリーデータが完成しました。", vbInformation, "MCS 2020 - Summarydata_Creation"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If
    
' 集計サマリーデータ１回作成処理
    fn_cnt = 1
    crs_cnt = 1
    tabinst_fn = Application.GetOpenFilename("集計設定ファイル,*.xlsx", , "集計設定ファイルを開く")
    If tabinst_fn = "False" Then
        ' キャンセルボタンの処理
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf tabinst_fn = "" Then
        MsgBox "集計する［集計設定ファイル］を選択してください。", vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ' フルパスからファイル名の取得
    tabinst_fn = Dir(tabinst_fn)

    wb.Activate
    ws_mainmenu.Select
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & tabinst_fn) <> "" Then
        Workbooks.Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & _
         ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & tabinst_fn
        Set wb_tabinst = Workbooks(tabinst_fn)
        Set ws_tabinst = wb_tabinst.ActiveSheet
        If ws_tabinst.Cells(2, 4) = "" Then
            MsgBox "集計設定ファイル［" & tabinst_fn & "］で、集計するデータファイルが入力されていません。", vbExclamation, "MCS 2020 - Summarydata_Creation"
            Application.StatusBar = False
            End
        End If
        outdata_fn = ws_tabinst.Cells(2, 4)
        weight_flg = ws_tabinst.Cells(2, 12)    ' ウエイト集計の有無フラグ取得
        If weight_flg = "" Then weight_flg = "なし"
    Else
        MsgBox "集計設定ファイル［" & tabinst_fn & "］が存在しません。", vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        End
    End If
    
    Call Outdata_Open
    Call Setup_Hold
    
    wb_outdata.Activate
    ws_outdata.Select
    outdata_row = Cells(Rows.Count, setup_col).End(xlUp).Row
    outdata_col = Cells(1, Columns.Count).End(xlToLeft).Column
    
    If weight_flg = "あり" Then   ' ウエイト（補正値）の取得
        w_index = Qcode_Match("weight")
    End If
    
    ' 2018/05/28 - サマリーの出力先は［SUM］フォルダ固定に変更。
'    summary_fd = Replace(Left(tabinst_fn, InStr(tabinst_fn, ".") - 1), ws_mainmenu.Cells(gcode_row, gcode_col), "")
    summary_fd = "SUM"
    summary_fn = Left(tabinst_fn, InStr(tabinst_fn, ".") - 1) & "_sum.xlsx"
    
'2018/04/26 - 追記 ==========================
    Open file_path & "\" & summary_fd & "\" & summary_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(summary_fn).Close
    End If
    If Dir(file_path & "\" & summary_fd & "\" & summary_fn) <> "" Then
        Kill file_path & "\" & summary_fd & "\" & summary_fn
    End If
'============================================
    
    Call Cross_Setting(crs_cnt, fn_cnt)
    
    ' データファイルのクローズ
    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFA:XFD").Select
    Selection.ClearContents
    Range("B7").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
    
    ' 集計設定ファイルのクローズ
    wb_tabinst.Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Set wb_tabinst = Nothing
    Set ws_tabinst = Nothing
    
'2020/01/10 - MCODE処理追記 =================
    Call Mcode_Setting
'============================================
    
    ' 集計サマリーファイルを保存してクローズ
    wb_summary.Activate
    ws_summary1.Select
    
    If Dir(file_path & "\" & summary_fd, vbDirectory) = "" Then
        MkDir file_path & "\" & summary_fd
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=file_path & "\" & summary_fd & "\" & summary_fn
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Set wb_summary = Nothing
    Set ws_summary0 = Nothing
    Set ws_summary1 = Nothing
    Set ws_summary2 = Nothing
    Set ws_summary3 = Nothing
    
    wb.Activate
    ws_setup.Select
    ws_setup.Cells(1, 1).Select
    ws_mainmenu.Select
    ws_mainmenu.Cells(3, 8).Select
    
' システムログの出力
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "23"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 23"
    End If
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
    Close #1
    If Err.Number > 0 Then
        Close #1
    End If
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his") = "" Then
        Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
         "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Output As #1
        Print #1, ws_mainmenu.Cells(gcode_row, gcode_col) & " MCS 2020 operation history"
        Close #1
    End If
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計サマリーデータの作成：対象ファイル［" & tabinst_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "集計サマリーデータが完成しました。", vbInformation, "MCS 2020 - Summarydata_Creation"
End Sub

Private Sub Cross_Setting(ByVal crs_cntx As Long, ByVal fn_cntx As Long)
' 集計サマリーファイルの作成
    Dim waitTime As Variant
    Dim i_cnt As Long
    Dim f_cnt As Long
    Dim max_row As Long
    Dim select_state As String
    
    c_cnt = 1
    sum_row = 1: sum_col = 1
    div_row = 1: div_col = 1
     
    ' 集計サマリー出力用ファイルの展開
    Workbooks.Add
    Worksheets.Add after:=Worksheets(1)
    Worksheets.Add after:=Worksheets(2)
    Worksheets.Add after:=Worksheets(3)
    
    Set wb_summary = ActiveWorkbook
    Set ws_summary0 = wb_summary.Worksheets("Sheet1")
    Set ws_summary1 = wb_summary.Worksheets("Sheet2")
    Set ws_summary2 = wb_summary.Worksheets("Sheet3")
    Set ws_summary3 = wb_summary.Worksheets("Sheet4")
    
    ws_summary0.Name = "目次"
    ws_summary1.Name = "Ｎ％表"
    ws_summary2.Name = "Ｎ表"
    ws_summary3.Name = "％表"
    
    ' 集計設定ファイルの情報をユーザー型変数に格納
    wb_tabinst.Activate
    ws_tabinst.Select
    
    max_row = ws_tabinst.Cells(Rows.Count, setup_col).End(xlUp).Row
    ReDim c_data(max_row)
    
'2018/05/01 - 追記 ==========================
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Cells(7, 1), Cells(max_row, 1)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range(Cells(6, 1), Cells(max_row, 21))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'============================================
    
    Application.ScreenUpdating = False
    Load UserForm52
    UserForm52.StartUpPosition = 1
    UserForm52.Show vbModeless
    UserForm52.Repaint
    Application.Visible = False
    AppActivate UserForm52.Caption
    UserForm52.Label2.Caption = tabinst_fn
    UserForm52.Label3.Caption = outdata_fn
    UserForm52.Label4.Caption = summary_fn
    UserForm52.Label5.Caption = "出力先：" & file_path & "\" & summary_fd
    UserForm52.Label7.Caption = "[" & crs_cntx & "/" & fn_cntx & "ファイル]"
    
    For i_cnt = 7 To max_row
        DoEvents
        UserForm52.Label1.Caption = Int(i_cnt / max_row * 100) & "%"
        UserForm52.Label6.Caption = "集計中" & Status_Dot(c_cnt)
        wb_tabinst.Activate
        ws_tabinst.Select
        If Mid(Cells(i_cnt, 1), 1, 1) <> "*" Then
            c_data(c_cnt).hyo_num = Cells(i_cnt, 1)
            c_data(c_cnt).f_code = Cells(i_cnt, 2)
            c_data(c_cnt).a_code = Cells(i_cnt, 3)
            c_data(c_cnt).r_code = Cells(i_cnt, 4)
            c_data(c_cnt).fna_flg = Cells(i_cnt, 5)
            c_data(c_cnt).ana_flg = Cells(i_cnt, 6)
            c_data(c_cnt).bosu_flg = Cells(i_cnt, 7)
            c_data(c_cnt).sum_flg = Cells(i_cnt, 8)
            c_data(c_cnt).ave_flg = Cells(i_cnt, 9)
            c_data(c_cnt).sd_flg = Cells(i_cnt, 10)
            c_data(c_cnt).min_flg = Cells(i_cnt, 11)
            c_data(c_cnt).q1_flg = Cells(i_cnt, 12)
            c_data(c_cnt).med_flg = Cells(i_cnt, 13)
            c_data(c_cnt).q3_flg = Cells(i_cnt, 14)
            c_data(c_cnt).max_flg = Cells(i_cnt, 15)
            c_data(c_cnt).mod_flg = Cells(i_cnt, 16)
            c_data(c_cnt).sel_code = Cells(i_cnt, 17)
            c_data(c_cnt).sel_value = Val(Cells(i_cnt, 18))
            c_data(c_cnt).ken_flg = Cells(i_cnt, 19)
            c_data(c_cnt).yuko_flg = Cells(i_cnt, 20)
            c_data(c_cnt).nobe_flg = Cells(i_cnt, 21)

            c_data(c_cnt).top1_flg = Cells(i_cnt, 22)
            c_data(c_cnt).sort_flg = Cells(i_cnt, 23)
            c_data(c_cnt).exct_flg = Cells(i_cnt, 24)
            c_data(c_cnt).graph_flg = Cells(i_cnt, 25)

            select_flg = 0
            select_state = "0000000"
            se_msg = ""
            fs1_msg = "": fs2_msg = "": fs3_msg = ""
            as1_msg = "": as2_msg = "": as3_msg = ""
            
' 集計設定の集計条件（セレクト）を取得
            se_index = 0
            
            ' QCODEをサーチ
            If c_data(c_cnt).sel_code <> "" Then
                select_flg = 1
                Mid(select_state, 1, 1) = "1"
                se_index = Qcode_Match(c_data(c_cnt).sel_code)
                
                ' マルチアンサーの処理
                se_cnt = 0
                If (q_data(se_index).q_format = "M") Or (Mid(q_data(se_index).q_format, 1, 1) = "L") Then
                    se_cnt = c_data(c_cnt).sel_value
                End If
                
                ' 表題とカテゴリーのコメントの設定がなかったら、セレクトコメントは非表示にする。
                If (q_data(se_index).q_title = "") And (q_data(se_index).q_ct(c_data(c_cnt).sel_value) = "") Then
                    se_msg = ""
                Else
                    se_msg = q_data(se_index).q_title & "：" & q_data(se_index).q_ct(c_data(c_cnt).sel_value)
                End If
            End If

' 表側の設定情報を取得
            f_index = 0
            fs1_index = 0: fs2_index = 0: fs3_index = 0
            
            ' QCODEをサーチ
            If c_data(c_cnt).f_code <> "" Then
                f_index = Qcode_Match(c_data(c_cnt).f_code)
                
                ' マルチアンサーの処理
                fma_cnt = 0
                If (q_data(f_index).q_format = "M") Or (Mid(q_data(se_index).q_format, 1, 1) = "L") Then
                    fma_cnt = q_data(f_index).ct_count
                End If
                
                ' セレクト①をサーチ
                If q_data(f_index).sel_code1 <> "" Then
                    select_flg = 1
                    Mid(select_state, 2, 1) = "1"
                    fs1_index = Qcode_Match(q_data(f_index).sel_code1)
                    
                    ' マルチアンサーの処理
                    fs1_cnt = 0
                    If (q_data(fs1_index).q_format = "M") Or (Mid(q_data(fs1_index).q_format, 1, 1) = "L") Then
                        fs1_cnt = q_data(f_index).sel_value1
                    End If
                    fs1_msg = q_data(fs1_index).q_title & "：" & q_data(fs1_index).q_ct(q_data(f_index).sel_value1)
                End If
                
                ' セレクト②をサーチ
                If q_data(f_index).sel_code2 <> "" Then
                    select_flg = 1
                    Mid(select_state, 3, 1) = "1"
                    fs2_index = Qcode_Match(q_data(f_index).sel_code2)
                    
                    ' マルチアンサーの処理
                    fs2_cnt = 0
                    If (q_data(fs2_index).q_format = "M") Or (Mid(q_data(fs2_index).q_format, 1, 1) = "L") Then
                        fs2_cnt = q_data(f_index).sel_value2
                    End If
                    fs2_msg = q_data(fs2_index).q_title & "：" & q_data(fs2_index).q_ct(q_data(f_index).sel_value2)
                End If
                
                ' セレクト③をサーチ
                If q_data(f_index).sel_code3 <> "" Then
                    select_flg = 1
                    Mid(select_state, 4, 1) = "1"
                    fs3_index = Qcode_Match(q_data(f_index).sel_code3)
                    
                    ' マルチアンサーの処理
                    fs3_cnt = 0
                    If (q_data(fs3_index).q_format = "M") Or (Mid(q_data(fs3_index).q_format, 1, 1) = "L") Then
                        fs3_cnt = q_data(f_index).sel_value3
                    End If
                    fs3_msg = q_data(fs3_index).q_title & "：" & q_data(fs3_index).q_ct(q_data(f_index).sel_value3)
                End If
            End If

' 表頭の設定情報を取得
            a_index = 0
            ra_index = 0
            as1_index = 0: as2_index = 0: as3_index = 0
                
            If c_data(c_cnt).a_code <> "" Then
                a_index = Qcode_Match(c_data(c_cnt).a_code)
                
                ' 実数QCODEをサーチ
                If c_data(c_cnt).r_code <> "" Then
                    ra_index = Qcode_Match(c_data(c_cnt).r_code)
                
                    ' 実数項目と表頭項目のセレクトのチェック
                    If q_data(a_index).sel_code1 <> q_data(ra_index).sel_code1 Then
                        MsgBox "表頭 QCODE［" & q_data(a_index).q_code & "］と" & vbCrLf & _
                        "実数 QCODE［" & q_data(ra_index).q_code & "］の" & vbCrLf & _
                        "条件①を同じ設定にしてください。" & vbCrLf & vbCrLf & _
                        "【TIPS】" & vbCrLf & "設定画面で、上記 QCODE の条件①の設定を確認してください。", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                    If q_data(a_index).sel_code2 <> q_data(ra_index).sel_code2 Then
                        MsgBox "表頭 QCODE［" & q_data(a_index).q_code & "］と" & vbCrLf & _
                        "実数 QCODE［" & q_data(ra_index).q_code & "］の" & vbCrLf & _
                        "条件②を同じ設定にしてください。" & vbCrLf & vbCrLf & _
                        "【TIPS】" & vbCrLf & "設定画面で、上記 QCODE の条件②の設定を確認してください。", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                    If q_data(a_index).sel_code3 <> q_data(ra_index).sel_code3 Then
                        MsgBox "表頭 QCODE［" & q_data(a_index).q_code & "］と" & vbCrLf & _
                        "実数 QCODE［" & q_data(ra_index).q_code & "］の" & vbCrLf & _
                        "条件③を同じ設定にしてください。" & vbCrLf & vbCrLf & _
                        "【TIPS】" & vbCrLf & "設定画面で、上記 QCODE の条件③の設定を確認してください。", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                End If
                
                ' マルチアンサーの処理
                ama_cnt = 0
                If (q_data(a_index).q_format = "M") Or (Mid(q_data(a_index).q_format, 1, 1) = "L") Then
                    ama_cnt = q_data(a_index).ct_count
                End If
                
                ' セレクト①をサーチ
                If q_data(a_index).sel_code1 <> "" Then
                    select_flg = 1
                    Mid(select_state, 5, 1) = "1"
                    as1_index = Qcode_Match(q_data(a_index).sel_code1)
                    
                    ' マルチアンサーの処理
                    as1_cnt = 0
                    If (q_data(as1_index).q_format = "M") Or (Mid(q_data(as1_index).q_format, 1, 1) = "L") Then
                        as1_cnt = q_data(a_index).sel_value1
                    End If
                    as1_msg = q_data(as1_index).q_title & "：" & q_data(as1_index).q_ct(q_data(a_index).sel_value1)
                End If
                
                ' セレクト②をサーチ
                If q_data(a_index).sel_code2 <> "" Then
                    select_flg = 1
                    Mid(select_state, 6, 1) = "1"
                    as2_index = Qcode_Match(q_data(a_index).sel_code2)
                    
                    ' マルチアンサーの処理
                    as2_cnt = 0
                    If (q_data(as2_index).q_format = "M") Or (Mid(q_data(as2_index).q_format, 1, 1) = "L") Then
                        as2_cnt = q_data(a_index).sel_value2
                    End If
                    as2_msg = q_data(as2_index).q_title & "：" & q_data(as2_index).q_ct(q_data(a_index).sel_value2)
                End If
                
                ' セレクト③をサーチ
                If q_data(a_index).sel_code3 <> "" Then
                    select_flg = 1
                    Mid(select_state, 7, 1) = "1"
                    as3_index = Qcode_Match(q_data(a_index).sel_code3)
                    
                    ' マルチアンサーの処理
                    as3_cnt = 0
                    If (q_data(as3_index).q_format = "M") Or (Mid(q_data(as3_index).q_format, 1, 1) = "L") Then
                        as3_cnt = q_data(a_index).sel_value3
                    End If
                    as3_msg = q_data(as3_index).q_title & "：" & q_data(as3_index).q_ct(q_data(a_index).sel_value3)
                End If
            Else
                MsgBox "表頭の［QCODE］は、必ず指定してください。", vbExclamation, "MCS 2020 - Cross_Setting"
                Call Files_Close
                End
            End If
            
            wb_summary.Activate
            Call Cross_Index                        ' 目次処理へ
            Call Cross_Header                       ' 集計表ヘッダーコメント処理へ
            
            If select_flg = 1 Then
                Call Select_Flag(select_state)      ' セレクト処理へ
            End If
            
            If weight_flg = "あり" Then
                Call weight_ra                      ' ウエイト集計・実数値算出処理へ
            End If
            
            If c_data(c_cnt).fna_flg <> "E" Then
                Call Simple_Summary                 ' 集計値算出・実数設問処理へ（単純集計）
            Else
                sum_row = sum_row - 2
                div_row = div_row - 1
            End If
            
            If c_data(c_cnt).f_code <> "" Then
                For f_cnt = 1 To q_data(f_index).ct_count
                    sum_row = sum_row + 2
                    div_row = div_row + 1
                    Call Cross_Tabulation(f_cnt)    ' 集計値算出・実数設問処理へ（クロス集計）
                Next f_cnt
                
                If c_data(c_cnt).fna_flg <> "N" Then
                    sum_row = sum_row + 2
                    div_row = div_row + 1
                    Call FaceNa_Tabulation          ' 集計値算出・実数設問処理へ（クロス集計）※表側無回答
                End If
            End If
            
            sum_row = sum_row + 3
            sum_col = 1
            
            div_row = div_row + 2
            div_col = 1
            
            c_cnt = c_cnt + 1
        End If
    Next i_cnt
    UserForm52.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")
    Application.Wait waitTime
    Application.Visible = True
    Unload UserForm52
End Sub

Private Sub Files_Close()
' エラー終了時のファイルクローズ
    Application.DisplayAlerts = False
    
    ' データファイル
    wb_outdata.Activate
    ActiveWorkbook.Close
    
    ' 集計設定ファイル
    wb_tabinst.Activate
    ActiveWorkbook.Close
    
    ' 集計サマリーファイル
    wb_summary.Activate
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    wb.Activate
    ws_setup.Select
    ws_setup.Cells(1, 1).Select
    ws_mainmenu.Select
    ws_mainmenu.Cells(3, 8).Select
    Application.StatusBar = False
    Application.Visible = True
    Unload UserForm52
End Sub

Private Sub Cross_Index()
    Dim sel_com As String
' 目次の作成
    If c_cnt = 1 Then
        ws_summary0.Cells(1, 1) = "連番"
        ws_summary0.Cells(1, 2) = "表№"
        ws_summary0.Cells(1, 3) = "MCODE"
        ws_summary0.Cells(1, 4) = "表側（縦軸）"
        ws_summary0.Cells(1, 5) = "表頭（横軸）"
        ws_summary0.Cells(1, 6) = "集計条件"
        ws_summary0.Cells(1, 7) = "リンク"
    
        If weight_flg = "あり" Then   ' ウエイト（補正値）の取得
            ws_summary0.Cells(1, 1) = "連番ウあ"
        End If
    End If
    
    ws_summary0.Cells(c_cnt + 1, 1) = c_cnt
    ws_summary0.Cells(c_cnt + 1, 2) = "'" & c_data(c_cnt).hyo_num
    ws_summary0.Cells(c_cnt + 1, 3) = q_data(a_index).m_code
    ws_summary0.Cells(c_cnt + 1, 4) = q_data(f_index).q_title
    ws_summary0.Cells(c_cnt + 1, 5) = q_data(a_index).q_title
    
    sel_com = ""
    If se_msg <> "" Then
        sel_com = sel_com & se_msg
    End If
    
    If fs1_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs1_msg
        Else
            sel_com = sel_com & vbLf & fs1_msg
        End If
    End If

    If fs2_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs2_msg
        Else
            sel_com = sel_com & vbLf & fs2_msg
        End If
    End If

    If fs3_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs3_msg
        Else
            sel_com = sel_com & vbLf & fs3_msg
        End If
    End If

    If as1_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as1_msg
        Else
            sel_com = sel_com & vbLf & as1_msg
        End If
    End If

    If as2_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as2_msg
        Else
            sel_com = sel_com & vbLf & as2_msg
        End If
    End If

    If as3_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as3_msg
        Else
            sel_com = sel_com & vbLf & as3_msg
        End If
    End If
    ws_summary0.Cells(c_cnt + 1, 6) = sel_com

    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 7), _
     Address:="", SubAddress:="'" & ws_summary1.Name & "'!A" & sum_row, TextToDisplay:="Ｎ％表"
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 8), _
     Address:="", SubAddress:="'" & ws_summary2.Name & "'!A" & div_row, TextToDisplay:="Ｎ表"
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 9), _
     Address:="", SubAddress:="'" & ws_summary3.Name & "'!A" & div_row, TextToDisplay:="％表"
End Sub

Private Sub Cross_Header()
    Dim i_cnt As Long
    Dim temp_row As Long
    Dim temp_col As Long
    Dim unit_cm As String
    Dim format_cm As String
    Dim bosu_cm As String
' 集計表ヘッダーコメント処理
    ' 表№の処理
    ' 下記はハイパーリンク文字列として出力するパターン
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary1.Cells(sum_row, sum_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!G" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary2.Cells(div_row, div_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!H" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary3.Cells(div_row, div_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!I" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    ' 下記はそのままテキストを文字列として出力するパターン
    'ws_summary1.Cells(sum_row, sum_col).Value = "'" & c_data(c_cnt).hyo_num
    'ws_summary2.Cells(div_row, div_col).Value = "'" & c_data(c_cnt).hyo_num
    'ws_summary3.Cells(div_row, div_col).Value = "'" & c_data(c_cnt).hyo_num
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' MCODEの処理
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).m_code
    ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).m_code
    ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).m_code
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    
    ' 表題の処理
    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).q_title
    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).q_title
    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).q_title
    sum_row = sum_row + 1
    div_row = div_row + 1
    
' セレクトコメントの出力順番
' 【１番目】集計設定セレクト
' 【２番目】表側セレクト①
' 【３番目】表側セレクト②
' 【４番目】表側セレクト③
' 【５番目】表頭セレクト①
' 【６番目】表頭セレクト②
' 【７番目】表頭セレクト③
'
    ' 集計設定セレクトコメントの処理
    If se_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【集計条件】" & se_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【集計条件】" & se_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【集計条件】" & se_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表側セレクト①コメントの処理
    If fs1_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表側集計条件】" & fs1_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表側集計条件】" & fs1_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表側集計条件】" & fs1_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表側セレクト②コメントの処理
    If fs2_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表側集計条件】" & fs2_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表側集計条件】" & fs2_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表側集計条件】" & fs2_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表側セレクト③コメントの処理
    If fs3_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表側集計条件】" & fs3_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表側集計条件】" & fs3_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表側集計条件】" & fs3_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表頭セレクト①コメントの処理
    If as1_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表頭集計条件】" & as1_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表頭集計条件】" & as1_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表頭集計条件】" & as1_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表頭セレクト②コメントの処理
    If as2_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表頭集計条件】" & as2_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表頭集計条件】" & as2_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表頭集計条件】" & as2_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' 表頭セレクト③コメントの処理
    If as3_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "【表頭集計条件】" & as3_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "【表頭集計条件】" & as3_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "【表頭集計条件】" & as3_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If

'---------------------------------
' 拡張オプションの処理 - 2018.9.14
    ' TOP1・マーキングフラグ
    ws_summary1.Cells(sum_row, sum_col - 2).Value = c_data(c_cnt).top1_flg
    ws_summary2.Cells(div_row, div_col - 2).Value = c_data(c_cnt).top1_flg
    ws_summary3.Cells(div_row, div_col - 2).Value = c_data(c_cnt).top1_flg

    ' CTソート・降順フラグ＆除外CTフラグ
    ws_summary1.Cells(sum_row + 1, sum_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg
    ws_summary2.Cells(div_row + 1, div_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg
    ws_summary3.Cells(div_row + 1, div_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg

    ' グラフ・作成フラグ
    ws_summary1.Cells(sum_row + 2, sum_col - 2).Value = c_data(c_cnt).graph_flg
    ws_summary2.Cells(div_row + 2, div_col - 2).Value = c_data(c_cnt).graph_flg
    ws_summary3.Cells(div_row + 2, div_col - 2).Value = c_data(c_cnt).graph_flg
'---------------------------------
    
    ' 表頭コメントの処理
    temp_row = sum_row
    temp_col = sum_col
    
    ' 表頭カテゴリー番号の処理
    sum_col = sum_col + 3
    div_col = div_col + 3
    For i_cnt = 1 To q_data(a_index).ct_count
        ws_summary1.Cells(sum_row, sum_col).Value = i_cnt
        ws_summary2.Cells(div_row, div_col).Value = i_cnt
        ws_summary3.Cells(div_row, div_col).Value = i_cnt
        sum_col = sum_col + 1
        div_col = div_col + 1
    Next i_cnt
    
    ' カテゴリー番号［無回答］の処理
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "N/A"
        ws_summary2.Cells(div_row, div_col).Value = "N/A"
        ws_summary3.Cells(div_row, div_col).Value = "N/A"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    sum_row = sum_row + 1
    sum_col = temp_col
    div_row = div_row + 1
    div_col = temp_col
    
    ' コメント［設問形式］と［構成比母数］の処理 - 2018.5.24 追加
    format_cm = ""
    If q_data(a_index).q_format = "S" Then
        format_cm = "［設問形式：単一回答］"
    ElseIf q_data(a_index).q_format = "M" Then
        format_cm = "［設問形式：複数回答］"
    ElseIf Mid(q_data(a_index).q_format, 1, 1) = "L" Then
        If q_data(a_index).ct_loop = 1 Then
            format_cm = "［設問形式：単一回答］"
        Else
            format_cm = "［設問形式：限定複数回答］"
        End If
    ElseIf q_data(a_index).q_format = "R" Then
        format_cm = "［設問形式：実数回答］"
    ElseIf q_data(a_index).q_format = "H" Then
        format_cm = "［設問形式：Ｈカーソル］"
    End If
    
    bosu_cm = ""
    If c_data(c_cnt).bosu_flg = "Y" Then
        bosu_cm = "［構成比母数：有効回答数］"
    Else
        bosu_cm = "［構成比母数：全体］"
    End If
    
    ws_summary1.Cells(sum_row, sum_col).Value = format_cm & vbCrLf & bosu_cm
    ws_summary2.Cells(div_row, div_col).Value = format_cm & vbCrLf & bosu_cm
    ws_summary3.Cells(div_row, div_col).Value = format_cm & vbCrLf & bosu_cm
    
    ' コメント［全体］の処理
    ws_summary1.Cells(sum_row, sum_col + 2).Value = "件数"
    ws_summary2.Cells(div_row, div_col + 2).Value = "件数"
    ws_summary3.Cells(div_row, div_col + 2).Value = "件数"
    sum_col = sum_col + 3
    div_col = div_col + 3
    
    ' コメント［表頭］の処理
    For i_cnt = 1 To q_data(a_index).ct_count
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).q_ct(i_cnt)
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).q_ct(i_cnt)
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).q_ct(i_cnt)
        sum_col = sum_col + 1
        div_col = div_col + 1
    Next i_cnt
    
    ' コメント［無回答］の処理
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "無回答"
        ws_summary2.Cells(div_row, div_col).Value = "無回答"
        ws_summary3.Cells(div_row, div_col).Value = "無回答"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［有効回答欄］の処理
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "有効回答数"
        ws_summary2.Cells(div_row, div_col).Value = "有効回答数"
        ws_summary3.Cells(div_row, div_col).Value = "有効回答数"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［延べ回答欄］の処理
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "延べ回答数"
        ws_summary2.Cells(div_row, div_col).Value = "延べ回答数"
        ws_summary3.Cells(div_row, div_col).Value = "延べ回答数"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［合計］の処理
    If c_data(c_cnt).sum_flg = "Y" Then
        If q_data(ra_index).r_unit <> "" Then
            unit_cm = "（" & q_data(ra_index).r_unit & "）"
        Else
            unit_cm = ""
        End If
        ws_summary1.Cells(sum_row, sum_col).Value = "合計" & unit_cm
        ws_summary2.Cells(div_row, div_col).Value = "合計" & unit_cm
        ws_summary3.Cells(div_row, div_col).Value = "合計" & unit_cm
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［平均］の処理
    If c_data(c_cnt).ave_flg <> "" Then
        If q_data(ra_index).r_unit <> "" Then
            unit_cm = "（" & q_data(ra_index).r_unit & "）"
        Else
            unit_cm = ""
        End If
        ws_summary1.Cells(sum_row, sum_col).Value = "平均" & unit_cm
        ws_summary2.Cells(div_row, div_col).Value = "平均" & unit_cm
        ws_summary3.Cells(div_row, div_col).Value = "平均" & unit_cm
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［標準偏差］の処理
    If c_data(c_cnt).sd_flg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "標準偏差"
        ws_summary2.Cells(div_row, div_col).Value = "標準偏差"
        ws_summary3.Cells(div_row, div_col).Value = "標準偏差"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［最小値］の処理
    If c_data(c_cnt).min_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "最小値"
        ws_summary2.Cells(div_row, div_col).Value = "最小値"
        ws_summary3.Cells(div_row, div_col).Value = "最小値"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［第１四分位］の処理
    If c_data(c_cnt).q1_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "第１四分位"
        ws_summary2.Cells(div_row, div_col).Value = "第１四分位"
        ws_summary3.Cells(div_row, div_col).Value = "第１四分位"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［中央値］の処理
    If c_data(c_cnt).med_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "中央値"
        ws_summary2.Cells(div_row, div_col).Value = "中央値"
        ws_summary3.Cells(div_row, div_col).Value = "中央値"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［第３四分位］の処理
    If c_data(c_cnt).q1_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "第３四分位"
        ws_summary2.Cells(div_row, div_col).Value = "第３四分位"
        ws_summary3.Cells(div_row, div_col).Value = "第３四分位"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［最大値］の処理
    If c_data(c_cnt).max_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "最大値"
        ws_summary2.Cells(div_row, div_col).Value = "最大値"
        ws_summary3.Cells(div_row, div_col).Value = "最大値"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' コメント［最頻値］の処理
    If c_data(c_cnt).mod_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "最頻値"
        ws_summary2.Cells(div_row, div_col).Value = "最頻値"
        ws_summary3.Cells(div_row, div_col).Value = "最頻値"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    sum_row = sum_row + 1
    sum_col = temp_col
    div_row = div_row + 1
    div_col = temp_col
End Sub

Private Sub Select_Flag(ByVal state_flag As Long)
    Dim i_cnt As Long
' 集計対象データのごとのセレクトフラグ加工
' 【概要】集計対象データとなるファイルの最終列［16384］列目に、セレクト条件を満たしていればフラグをたてる。
'
    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFD:XFD").Select
    Selection.ClearContents
    Cells(6, 16384) = "Select"
    
    ' サンプルベースでのセレクトの有無判定
    For i_cnt = 7 To outdata_row
        Select Case state_flag
        Case "0000000"
            '処理なし（セレクトなし）
        Case "1111111"
            Call Select_Decision01(i_cnt)
        Case "1111110"
            Call Select_Decision02(i_cnt)
        Case "1111100"
            Call Select_Decision03(i_cnt)
        Case "1111000"
            Call Select_Decision04(i_cnt)
        Case "1110000"
            Call Select_Decision05(i_cnt)
        Case "1100000"
            Call Select_Decision06(i_cnt)
        Case "1000000"
            Call Select_Decision07(i_cnt)
        Case "1110111"
            Call Select_Decision08(i_cnt)
        Case "1100111"
            Call Select_Decision09(i_cnt)
        Case "1000111"
            Call Select_Decision10(i_cnt)
        Case "1110110"
            Call Select_Decision11(i_cnt)
        Case "1110100"
            Call Select_Decision12(i_cnt)
        Case "1100110"
            Call Select_Decision13(i_cnt)
        Case "1100100"
            Call Select_Decision14(i_cnt)
        Case "1000110"
            Call Select_Decision15(i_cnt)
        Case "1000100"
            Call Select_Decision16(i_cnt)
        Case "0111111"
            Call Select_Decision17(i_cnt)
        Case "0111110"
            Call Select_Decision18(i_cnt)
        Case "0111100"
            Call Select_Decision19(i_cnt)
        Case "0111000"
            Call Select_Decision20(i_cnt)
        Case "0110000"
            Call Select_Decision21(i_cnt)
        Case "0100000"
            Call Select_Decision22(i_cnt)
        Case "0110111"
            Call Select_Decision23(i_cnt)
        Case "0100111"
            Call Select_Decision24(i_cnt)
        Case "0000111"
            Call Select_Decision25(i_cnt)
        Case "0110110"
            Call Select_Decision26(i_cnt)
        Case "0110100"
            Call Select_Decision27(i_cnt)
        Case "0100110"
            Call Select_Decision28(i_cnt)
        Case "0100100"
            Call Select_Decision29(i_cnt)
        Case "0000110"
            Call Select_Decision30(i_cnt)
        Case "0000100"
            Call Select_Decision31(i_cnt)
        End Select
    Next i_cnt
End Sub

Private Sub Select_Decision01(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表側③・表頭①・表頭②・表頭③
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision02(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表側③・表頭①・表頭②
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision03(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表側③・表頭①
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision04(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表側③
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision05(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision06(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision07(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Select_Decision08(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表頭①・表頭②・表頭③
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision09(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表頭①・表頭②・表頭③
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision10(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表頭①・表頭②・表頭③
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision11(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表頭①・表頭②
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision12(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表側②・表頭①
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision13(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表頭①・表頭②
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision14(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表側①・表頭①
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision15(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表頭①・表頭②
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision16(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 集計設定・表頭①
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision17(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表側③・表頭①・表頭②・表頭③
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision18(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表側③・表頭①・表頭②
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision19(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表側③・表頭①
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision20(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表側③
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision21(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision22(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Select_Decision23(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表頭①・表頭②・表頭③
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision24(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表頭①・表頭②・表頭③
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision25(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表頭①・表頭②・表頭③
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision26(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表頭①・表頭②
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision27(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表側②・表頭①
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision28(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表頭①・表頭②
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision29(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表側①・表頭①
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision30(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表頭①・表頭②
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision31(ByVal ix_cnt As Long)
' セレクト判定
' それぞれのカウンタ［*_cnt］は、該当項目のＳＡ（*_cnt=0）・ＭＡ（*_cnt<>0）判定です。
' 表頭①
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Simple_Summary()
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim gt_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
    Dim decimal_places As Integer
' 集計値算出・実数設問の処理（単純集計）
    On Error Resume Next
    temp_col = sum_col
    
    ' ［全体］の算出
    gt_cnt = 0
    If select_flg = 1 Then
        If weight_flg = "なし" Then
            gt_cnt = Application.WorksheetFunction. _
             CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
             Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
        Else
            gt_cnt = Application.WorksheetFunction. _
             SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
             Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), "<>", _
             Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
        End If
    Else
        If weight_flg = "なし" Then
            gt_cnt = Application.WorksheetFunction. _
             CountIf(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>" & "")
        Else
            gt_cnt = Application.WorksheetFunction. _
             SumIf(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), "<>", _
             Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
        End If
    End If
    
    ' ［無回答］の算出
    na_cnt = 0
    If ama_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "なし" Then
                na_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                na_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "なし" Then
                na_cnt = Application.WorksheetFunction. _
                 CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "")
            Else
                na_cnt = Application.WorksheetFunction. _
                 SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(a_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            na_cnt = na_cnt + 1
                        End If
                    End If
                Else
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(a_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                        End If
                    End If
                End If
            Else
                If weight_flg = "なし" Then
                    For ma_cnt = 1 To q_data(a_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        na_cnt = na_cnt + 1
                    End If
                Else
                    For ma_cnt = 1 To q_data(a_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' ［有効回答数］の算出
    vr_cnt = gt_cnt - na_cnt
    
    ' 件数［全体］の処理
    ws_summary1.Cells(sum_row, sum_col - 1).Value = "0"
    ws_summary2.Cells(div_row, div_col - 1).Value = "0"
    ws_summary3.Cells(div_row, div_col - 1).Value = "0"
    ws_summary1.Cells(sum_row, sum_col).Value = "　全　体"
    ws_summary2.Cells(div_row, div_col).Value = "　全　体"
    ws_summary3.Cells(div_row, div_col).Value = "　全　体"
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' 件数欄に［有効回答数］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' 件数欄に［全体］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = gt_cnt
        ws_summary2.Cells(div_row, div_col).Value = gt_cnt
        ws_summary3.Cells(div_row, div_col).Value = gt_cnt
    End If
    
    ' ウエイト集計時［全体］のセル書式設定
    If weight_flg = "あり" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' 件数［カテゴリー］の処理
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            Else
                If weight_flg = "なし" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If gt_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            Else
                If weight_flg = "なし" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If gt_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' 件数［無回答］の処理
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' ウエイト集計時［無回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If gt_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / gt_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / gt_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［有効回答］の処理
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' ウエイト集計時［有効回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If gt_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / gt_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / gt_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［延べ回答］の処理
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' ウエイト集計時［延べ回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If gt_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / gt_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / gt_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 集計設定ファイルの条件判定［実数設問用］
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If select_flg = 1 Then
        Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    End If
    
    If weight_flg = "なし" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' 各実数設問の処理へ
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' 各実数設問（ウエイト集計）の処理へ
    End If
    
    
    'オートフィルタの解除
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub Cross_Tabulation(ByVal x_cnt As Long)
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim face_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
    Dim decimal_places As Integer
' 集計値算出・実数設問の処理（クロス集計）
    On Error Resume Next
    temp_col = sum_col

    ' ［表側項目］の算出
    face_cnt = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            End If
        End If
    Else
        If select_flg = 1 Then
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            End If
        End If
    End If
    
    ' ［無回答］の算出
    na_cnt = 0
    If ama_cnt = 0 Then
        If fma_cnt = 0 Then
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                End If
            End If
        Else
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                End If
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If fma_cnt = 0 Then
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column) = x_cnt Then
                    If select_flg = 1 Then
                        If weight_flg = "なし" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "なし" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column + x_cnt - 1) = 1 Then
                    If select_flg = 1 Then
                        If weight_flg = "なし" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "なし" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' ［有効回答数］の算出
    vr_cnt = face_cnt - na_cnt
    
    ' 表側ヘッダーコメントの処理
    If x_cnt = 1 Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = q_data(f_index).q_title
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = q_data(f_index).q_title
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = q_data(f_index).q_title
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' 件数［表側項目］の処理
    ws_summary1.Cells(sum_row, sum_col - 2).Value = x_cnt
    ws_summary2.Cells(div_row, div_col - 2).Value = x_cnt
    ws_summary3.Cells(div_row, div_col - 2).Value = x_cnt
    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(f_index).q_ct(x_cnt)
    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary2.Cells(div_row, div_col).Value = q_data(f_index).q_ct(x_cnt)
    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary3.Cells(div_row, div_col).Value = q_data(f_index).q_ct(x_cnt)
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' 件数欄に［有効回答数］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' 件数欄に［表側項目全数］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = face_cnt
        ws_summary2.Cells(div_row, div_col).Value = face_cnt
        ws_summary3.Cells(div_row, div_col).Value = face_cnt
    End If
    
    ' ウエイト集計時［表側項目全数］のセル書式設定
    If weight_flg = "あり" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' 件数［カテゴリー］の処理
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' 件数［無回答］の処理
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' ウエイト集計時［無回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If face_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［有効回答］の処理
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' ウエイト集計時［有効回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［延べ回答］の処理
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' ウエイト集計時［延べ回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 集計設定ファイルの条件判定［実数設問用］
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, x_cnt, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, x_cnt, visibledropdown:=False
        End If
    Else
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column + x_cnt - 1, 1, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column + x_cnt - 1, 1, visibledropdown:=False
        End If
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    End If
    
    If weight_flg = "なし" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' 各実数設問の処理へ
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' 各実数設問（ウエイト集計）の処理へ
    End If
    
    'オートフィルタの解除
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub FaceNa_Tabulation()
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim face_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
' 集計値算出・実数設問の処理（クロス集計）※表側無回答
' 【概要】ファイルの［16383］列目に、表側無回答フラグをたてる。
'
    On Error Resume Next
    temp_col = sum_col

    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFC:XFC").Select
    Selection.ClearContents
    Cells(6, 16383) = "FMA[N/A]"

    ' ［表側無回答項目］の算出
    face_cnt = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "なし" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(f_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            face_cnt = face_cnt + 1
                            ws_outdata.Cells(i_cnt, 16383) = 1
                        End If
                    End If
                Else
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(f_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            face_cnt = face_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            ws_outdata.Cells(i_cnt, 16383) = 1
                        End If
                    End If
                End If
            Else
                If weight_flg = "なし" Then
                    For ma_cnt = 1 To q_data(f_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        face_cnt = face_cnt + 1
                        ws_outdata.Cells(i_cnt, 16383) = 1
                    End If
                Else
                    For ma_cnt = 1 To q_data(f_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        face_cnt = face_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                        ws_outdata.Cells(i_cnt, 16383) = 1
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' ［無回答］の算出
    na_cnt = 0
    If ama_cnt = 0 Then
        If fma_cnt = 0 Then
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                   na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                   na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                End If
            End If
        Else
            If select_flg = 1 Then
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "なし" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                End If
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If fma_cnt = 0 Then
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column) = "" Then
                    If select_flg = 1 Then
                        If weight_flg = "なし" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "なし" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(i_cnt, 16383) = 1 Then
                    If select_flg = 1 Then
                        If weight_flg = "なし" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "なし" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' ［有効回答数］の算出
    vr_cnt = face_cnt - na_cnt
    
    ' 件数［表側項目］の処理
    ws_summary1.Cells(sum_row, sum_col - 1).Value = "N/A"
    ws_summary2.Cells(div_row, div_col - 1).Value = "N/A"
    ws_summary3.Cells(div_row, div_col - 1).Value = "N/A"
    ws_summary1.Cells(sum_row, sum_col + 1).Value = "無回答"
    ws_summary2.Cells(div_row, div_col + 1).Value = "無回答"
    ws_summary3.Cells(div_row, div_col + 1).Value = "無回答"
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' 件数欄に［有効回答数］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' 件数欄に［表側項目全数］を出力
        ws_summary1.Cells(sum_row, sum_col).Value = face_cnt
        ws_summary2.Cells(div_row, div_col).Value = face_cnt
        ws_summary3.Cells(div_row, div_col).Value = face_cnt
    End If
    
    ' ウエイト集計時［全体］のセル書式設定
    If weight_flg = "あり" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' 件数［カテゴリー］の処理
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                ' 表側ＭＡのＮＡ処理＠表頭ＳＡ
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                ' 表側ＭＡのＮＡ処理＠表頭ＭＡ
                If select_flg = 1 Then
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "なし" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '構成比の算出
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' 件数［無回答］の処理
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' ウエイト集計時［無回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If face_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［有効回答］の処理
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' ウエイト集計時［有効回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 件数［延べ回答］の処理
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' ウエイト集計時［延べ回答］のセル書式設定
        If weight_flg = "あり" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 集計設定ファイルの条件判定［実数設問用］
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, "=", visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, "=", visibledropdown:=False
        End If
    Else
        ' 表側ＭＡのＮＡ処理＠実数設問
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16383, 1, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16383, 1, visibledropdown:=False
        End If
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    Else
        filter_flg = 99
    End If
    
    If weight_flg = "なし" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' 各実数設問の処理へ
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' 各実数設問（ウエイト集計）の処理へ
    End If
    
    'オートフィルタの解除
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub Real_Answer(ByVal f_flag As Integer, ByVal v_cnt As Long)
    On Error Resume Next
    Dim decimal_places As Integer
' 各実数設問の処理
    ' 実数設問［合計］の処理
    If c_data(c_cnt).sum_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                   
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［平均］の処理
    decimal_places = Val(c_data(c_cnt).ave_flg)
    If c_data(c_cnt).ave_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［標準偏差］の処理
    decimal_places = Val(c_data(c_cnt).sd_flg)
    If c_data(c_cnt).sd_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最小値］の処理
    If c_data(c_cnt).min_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［第１四分位］の処理
    If c_data(c_cnt).q1_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［中央値］の処理
    If c_data(c_cnt).med_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［第３四分位］の処理
    If c_data(c_cnt).q3_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最大値］の処理
    If c_data(c_cnt).max_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最頻値］の処理
    If c_data(c_cnt).mod_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
End Sub

Private Sub Real_Answer_WGT(ByVal f_flag As Integer, ByVal v_cnt As Long)
    On Error Resume Next
    Dim decimal_places As Integer
' 各実数設問（ウエイト集計）の処理
    ' 実数設問［合計］の処理
    If c_data(c_cnt).sum_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［平均］の処理
    decimal_places = Val(c_data(c_cnt).ave_flg)
    If c_data(c_cnt).ave_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［標準偏差］の処理
    decimal_places = Val(c_data(c_cnt).sd_flg)
    If c_data(c_cnt).sd_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最小値］の処理
    If c_data(c_cnt).min_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［第１四分位］の処理
    If c_data(c_cnt).q1_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
            Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［中央値］の処理
    If c_data(c_cnt).med_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［第３四分位］の処理
    If c_data(c_cnt).q3_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最大値］の処理
    If c_data(c_cnt).max_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            If v_cnt = 0 Then    ' 有効回答ゼロ判定
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' 実数設問［最頻値］の処理
    If c_data(c_cnt).mod_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' 表側無回答専用処理
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
End Sub

Private Sub weight_ra()
    Dim i_cnt As Long
' 集計対象データごとのウエイト集計用実数値の算出
' 【概要】集計対象データとなるファイルの［16382］列目に、実数値とウエイト（補正値）の積を出力。
'
    If ra_index <> 0 Then    ' 表頭項目に実数設問の指定があれば処理する。
        wb_outdata.Activate
        ws_outdata.Select
        Columns("XFB:XFB").Select
        Selection.ClearContents
        Cells(6, 16382) = "weight_ra"
    
        ' サンプルベースでの実数値×ウエイト（補正値）の算出
        For i_cnt = 7 To outdata_row
            If ws_outdata.Cells(i_cnt, q_data(ra_index).data_column) <> "" Then
                ws_outdata.Cells(i_cnt, 16382) = _
                 ws_outdata.Cells(i_cnt, q_data(ra_index).data_column) * ws_outdata.Cells(i_cnt, q_data(w_index).data_column)
            End If
        Next i_cnt
    End If
End Sub

Private Sub Mcode_Setting()
' MCODE処理 - 2020.1.10 追加、2020.3.26 編集
    Dim i_cnt As Long, m_cnt As Long
    Dim max_row As Long
    Dim s_pos As Long
    Dim m_code As String
    Dim hyo_num As String
    Dim bgn_row As Long, fin_row As Long
    Dim head_cm As String, face_cm As String
    
'【Ｎ％表】
    wb_summary.Activate
    ws_summary1.Select
    
    ' サマリーファイルの最終行取得（G列で取得してます）
    max_row = ws_summary1.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODEの検索（B列だけではなく、表番号とあわせて検索）
        If ws_summary1.Cells(i_cnt, 1) <> "" Then
            If ws_summary1.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "／")
                    head_cm = Left(ws_summary1.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary1.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' 単純集計の判定 - 2020.3.26
                    If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                      ws_summary1.Cells(i_cnt, 4) = head_cm
                      ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary1.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary1.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary1.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "／")
                        face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary1.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' 単純集計の判定 - 2020.3.26
                        If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                          ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary1.Cells(ActiveCell.Row, 4).Clear
                          ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary1.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary1.Cells(i_cnt, 2)
                        m_cnt = 1
                    
                        s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "／")
                        head_cm = Left(ws_summary1.Cells(i_cnt, 4), s_pos - 1)
                        face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary1.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' 単純集計の判定 - 2020.3.26
                        If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                          ws_summary1.Cells(i_cnt, 4) = head_cm
                          ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary1.Cells(ActiveCell.Row, 4) = head_cm
                          ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                          m_code = ws_summary1.Cells(i_cnt, 2)
                          m_cnt = m_cnt + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i_cnt

'【Ｎ表】
    wb_summary.Activate
    ws_summary2.Select
    
    ' サマリーファイルの最終行取得（G列で取得してます）
    max_row = ws_summary2.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODEの検索（B列だけではなく、表番号とあわせて検索）
        If ws_summary2.Cells(i_cnt, 1) <> "" Then
            If ws_summary2.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary2.Cells(i_cnt, 4), "／")
                    head_cm = Left(ws_summary2.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary2.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary2.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' 単純集計の判定 - 2020.3.26
                    If ws_summary2.Cells(ActiveCell.Row + 1, 3) = "" Then
                      ws_summary2.Cells(i_cnt, 4) = head_cm
                      ws_summary2.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary2.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary2.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary2.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary2.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary2.Cells(i_cnt, 4), "／")
                        face_cm = Mid(ws_summary2.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary2.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' 単純集計の判定 - 2020.3.26
                        If ws_summary2.Cells(ActiveCell.Row + 1, 3) = "" Then
                          ws_summary2.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary2.Cells(ActiveCell.Row, 4).Clear
                          ws_summary2.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary2.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary2.Cells(i_cnt, 2)
                        m_cnt = 1
                    End If
                End If
            End If
        End If
    Next i_cnt

'【％表】
    wb_summary.Activate
    ws_summary3.Select
    
    ' サマリーファイルの最終行取得（G列で取得してます）
    max_row = ws_summary3.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODEの検索（B列だけではなく、表番号とあわせて検索）
        If ws_summary3.Cells(i_cnt, 1) <> "" Then
            If ws_summary3.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary3.Cells(i_cnt, 4), "／")
                    head_cm = Left(ws_summary3.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary3.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary3.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' 単純集計の判定 - 2020.3.26
                    If ws_summary3.Cells(ActiveCell.Row + 1, 3) = "" Then
                      ws_summary3.Cells(i_cnt, 4) = head_cm
                      ws_summary3.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary3.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary3.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary3.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary3.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary3.Cells(i_cnt, 4), "／")
                        face_cm = Mid(ws_summary3.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary3.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' 単純集計の判定 - 2020.3.26
                        If ws_summary3.Cells(ActiveCell.Row + 1, 3) = "" Then
                          ws_summary3.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary3.Cells(ActiveCell.Row, 4).Clear
                          ws_summary3.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary3.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary3.Cells(i_cnt, 2)
                        m_cnt = 1
                    End If
                End If
            End If
        End If
    Next i_cnt

'【目次】
    wb_summary.Activate
    ws_summary0.Select
    
    ' サマリーファイルの最終行取得（A列で取得してます）
    max_row = ws_summary0.Cells(Rows.Count, 1).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 2 To max_row    ' ヘッダーがあるので開始は[2]から。
        ' MCODEの検索（B列だけではなく、表番号とあわせて検索）
        If ws_summary0.Cells(i_cnt, 2) <> "" Then
            If ws_summary0.Cells(i_cnt, 3) <> "" Then
                If ws_summary0.Cells(i_cnt, 4) = "" Then
                    If m_cnt = 1 Then
                        hyo_num = ws_summary0.Cells(i_cnt, 2)
                        m_code = ws_summary0.Cells(i_cnt, 3)
                        m_cnt = m_cnt + 1
                    Else
                        If ws_summary0.Cells(i_cnt, 3) = m_code Then
                            ws_summary0.Cells(i_cnt, 2) = "'" & hyo_num
                            m_cnt = m_cnt + 1
                        Else
                            hyo_num = ws_summary0.Cells(i_cnt, 2)
                            m_code = ws_summary0.Cells(i_cnt, 3)
                            m_cnt = 1
                        End If
                    End If
                End If
            End If
        End If
    Next i_cnt
End Sub
