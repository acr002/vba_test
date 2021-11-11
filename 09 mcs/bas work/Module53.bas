Attribute VB_Name = "Module53"
Option Explicit
    Dim mcs_ini(10) As String
    Dim ini_cnt As Integer
    Dim wb_spread As Workbook
    Dim ws_spread0 As Worksheet
    Dim ws_spread1 As Worksheet
    Dim ws_spread2 As Worksheet
    Dim ws_spread3 As Worksheet
    Dim summary_fd As String
    Dim summary_fn As String
    Dim spread_fn As String
    Dim hyo_cnt As Long
    Dim max_row As Long
    Dim np_max_row As Long
    Dim face_flag As Integer
    Dim i_cnt As Long

Sub Spreadsheet_Creation()
    Dim waitTime As Variant
    Dim yen_pos As Long
'2018/05/28 - 追記 ==========================
    Dim r_code As Integer
    Dim sum_tab() As String
    Dim sum_file As String
    Dim sum_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long
'2019/12/10 - 追記 ==========================
    Dim wgt_row As Long
'--------------------------------------------------------------------------------------------------'
'　集計表Excelファイルの作成 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.05.24　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    
    ' 設定ファイルからの情報取得
    ' (1)パス、(2)日本語フォント、(3)日本語フォントサイズ、(4)英数字フォント、(5)英数字フォントサイズ、(6)全体欄カラー、(7)罫線カラー
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
        Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
         "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Input As #1
        ini_cnt = 1
        Do Until EOF(1)
            DoEvents
            Line Input #1, mcs_ini(ini_cnt)
            Select Case ini_cnt
            Case 2
                If Mid(mcs_ini(ini_cnt), 1, 7) <> "J-FONT=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［FONT］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 3
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "J-FONT-SIZE=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［FONT-SIZE］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 4
                If Mid(mcs_ini(ini_cnt), 1, 7) <> "E-FONT=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［FONT］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 5
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "E-FONT-SIZE=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［FONT-SIZE］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 6
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "TOTAL-COLOR=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［TOTAL-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 16, 1) <> "," Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［TOTAL-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 20, 1) <> "," Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［TOTAL-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 7
                If Mid(mcs_ini(ini_cnt), 1, 13) <> "BORDER-COLOR=" Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［BORDER-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 17, 1) <> "," Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［BORDER-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 21, 1) <> "," Then
                    MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］の［BORDER-COLOR］設定を確認してください。" _
                     & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case Else
            End Select
            ini_cnt = ini_cnt + 1
        Loop
        Close #1
    Else
        MsgBox "設定ファイル［" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini］が見つかりません。" _
         & vbCrLf & "初期設定をもう一度行ってください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        Call Finishing_Mcs2017
        End
    End If
    
    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    ' SUMフォルダ内の*_sum.xlsx形式のファイル数をカウント
    sum_cnt = 0
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    Do Until sum_file = ""
        DoEvents
        sum_cnt = sum_cnt + 1
        sum_file = Dir()
    Loop
    
    ' SUMフォルダ内の*_sum.xlsx形式のファイル名を配列にセット
    ReDim sum_tab(sum_cnt)
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    For fn_cnt = 1 To sum_cnt
        DoEvents
        sum_tab(fn_cnt) = sum_file
        sum_file = Dir()
    Next fn_cnt
    fn_cnt = sum_cnt

' 集計表Excelファイル複数作成処理
    If sum_cnt > 0 Then
        r_code = MsgBox("SUMフォルダ内にある" & fn_cnt & "個の集計サマリーデータから、" & vbCrLf & "一括して集計表Excelファイルを作成しますか。" _
         & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "SUMフォルダ内の［*_sum.xlsx形式］のファイル数を" & vbCrLf & "表示しています。" _
         & vbCrLf & "「はい」　→ 集計サマリーデータを一括処理" & vbCrLf & "「いいえ」→ 集計サマリーファイルを選択してから処理", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Spreadsheet_Creation")
        If r_code = vbYes Then
            sum_cnt = 1
            For n_cnt = 1 To fn_cnt
                DoEvents
                wb.Activate
                ws_mainmenu.Select
                summary_fn = sum_tab(n_cnt)
                
                Open file_path & "\SUM\" & summary_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(summary_fn).Close
                Else
                    Workbooks.Open summary_fn
                End If
                
                Set wb_spread = Workbooks(summary_fn)
                Set ws_spread0 = wb_spread.Worksheets(1)
                Set ws_spread1 = wb_spread.Worksheets(2)
                Set ws_spread2 = wb_spread.Worksheets(3)
                Set ws_spread3 = wb_spread.Worksheets(4)
                
                wb_spread.Activate
                Sheets(Array("Ｎ％表", "Ｎ表", "％表")).Select
                Cells.Select
                
                ' 集計表シート全体のスタイル設定
                ActiveWindow.DisplayGridlines = False
                With Selection.Font
                    .Name = Mid(mcs_ini(2), 8)      ' 日本語フォントをセット
                    .Size = Mid(mcs_ini(3), 13)     ' 日本語フォントサイズをセット
                End With
                
                ' 集計表シート全体の表№のスタイル設定
                Columns(1).Select
                With Selection
                    .ColumnWidth = 4.88
                End With
                
                ' 集計表シート全体の表側カテゴリー番号のスタイル設定
                Columns(3).Select
                With Selection
                    .HorizontalAlignment = xlRight
                    .ColumnWidth = 4.25
                    .Font.Color = RGB(0, 112, 192)
                    .Font.Italic = True
                    .Font.Size = 7
                End With
                
                Range("A1").Select
                ws_spread0.Select
                ws_spread0.PageSetup.Orientation = xlLandscape
                hyo_cnt = ws_spread0.Cells(Rows.Count, setup_col).End(xlUp).Row - 1
                
                face_flag = 0       ' クロス集計表の有無判定用フラグ
                Application.ScreenUpdating = False
                Load Form_Progress
                Form_Progress.StartUpPosition = 1
                Form_Progress.Show vbModeless
                Form_Progress.Caption = "MCS 2020 - 集計表Excelファイルの作成"
                Form_Progress.Repaint
                progress_msg = "集計表Excelファイルの作成をキャンセルしました。"
                Application.Visible = False
                AppActivate Form_Progress.Caption
                
                ' Ｎ％表の処理
                ws_spread1.Select
                ws_spread1.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP1/5 集計表Excelファイル（Ｎ％表）作成中" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "ファイル]"
                    Call int_spreadsheet
                Next i_cnt
    
                ws_spread1.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "游ゴシック"
                    .Size = 8
                End With
                ws_spread1.Columns(2).Hidden = True     ' MCODE列の非表示
                Range("A1").Select
                
                ' Ｎ表の処理
                ws_spread2.Select
                ws_spread2.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
'---------------------------------
                np_max_row = ActiveCell.Row
'---------------------------------
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP2/5 集計表Excelファイル（Ｎ表）作成中" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "ファイル]"
                    Call ken_spreadsheet
                Next i_cnt
                
                ws_spread2.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "游ゴシック"
                    .Size = 8
                End With
                ws_spread2.Columns(2).Hidden = True     ' MCODE列の非表示
                Range("A1").Select
                
                ' ％表の処理
                ws_spread3.Select
                ws_spread3.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
'---------------------------------
                np_max_row = ActiveCell.Row
'---------------------------------
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP3/5 集計表Excelファイル（％表）作成中" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "ファイル]"
                    Call per_spreadsheet
                Next i_cnt
                
                ws_spread3.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "游ゴシック"
                    .Size = 8
                End With
                ws_spread3.Columns(2).Hidden = True     ' MCODE列の非表示
                Range("A1").Select
                
                Form_Progress.Label1.Caption = "100%"
                DoEvents
                Form_Progress.Label2.Caption = "STEP4/5 集計表Excelファイル（目次）作成中..."
                Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "ファイル]"
                waitTime = Now + TimeValue("0:00:01")
                Application.Wait waitTime
                
                ' 単純集計時の表側幅の調整
                If face_flag = 0 Then
                    wb_spread.Activate
                    Sheets(Array("Ｎ％表", "Ｎ表", "％表")).Select
                    ws_spread1.Columns(5).ColumnWidth = 12.37
                    ws_spread2.Columns(5).ColumnWidth = 12.37
                    ws_spread3.Columns(5).ColumnWidth = 12.37
                    Range("A1").Select
                End If
                
                ' 目次の処理
                wb_spread.Activate
                ws_spread0.Select
                ActiveWindow.DisplayGridlines = False
                Cells.Select
                With Selection.Font
                    .Name = Mid(mcs_ini(2), 8)      ' 日本語フォントをセット
                    .Size = 9                       ' フォントサイズをセット
                End With
                
                Columns(4).Select
                Columns(4).WrapText = True
                Columns(4).ColumnWidth = 42
                
                Columns(5).Select
                Columns(5).WrapText = True
                Columns(5).ColumnWidth = 42
                
                If WorksheetFunction.CountA(ws_spread0.Columns(6)) > 1 Then
                    Columns(6).Select
                    Columns(6).WrapText = True
                    Columns(6).ColumnWidth = 42
                End If
                ws_spread0.Columns(3).Hidden = True     ' MCODE列の非表示
                
                ' 目次の罫線処理
                Range("A1").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.LineStyle = xlContinuous
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                
                With Range(Selection, ActiveCell.SpecialCells(xlLastCell))
                    .Borders(xlInsideHorizontal).Weight = xlHairline
                End With
                With Range("A1:I1")
                    .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
                    .Borders(xlEdgeBottom).Weight = xlThin
                End With
                Range("G1:I1").Merge
                Range("G1").HorizontalAlignment = xlLeft
                Range("A2").Select
                ActiveWindow.FreezePanes = True
                Range("A1").Select
                
                ' ウエイトバック集計時の説明文
                If ws_spread0.Cells(1, 1) = "連番ウあ" Then
                    ws_spread0.Cells(1, 1) = "連番"
                    Range("A1").End(xlDown).Select
                    wgt_row = ActiveCell.Row
                    ws_spread0.Cells(wgt_row + 1, 1) = "※ウエイトバック集計を行っているため、計算過程で小数点が生じますが、本集計表上の数値は四捨五入して整数表記しています。"
                    Range("A1").Select
                End If
                
                ' 各シートの色を設定
                Sheets("Ｎ％表").Select
                With ActiveWorkbook.Sheets("Ｎ％表").Tab
                    .Color = 10066431
                    .TintAndShade = 0
                End With
                Sheets("Ｎ表").Select
                With ActiveWorkbook.Sheets("Ｎ表").Tab
                    .Color = 10092441
                    .TintAndShade = 0
                End With
                Sheets("％表").Select
                With ActiveWorkbook.Sheets("％表").Tab
                    .Color = 16764057
                    .TintAndShade = 0
                End With
                Sheets("目次").Select
                
                ' TOP1・GTソート処理 - 2018.10.04 追加
                ws_spread1.Select
                ws_spread1.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP5/5 集計表Excelファイル 最終調整中" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "ファイル]"
                    Call top1_sort
                Next i_cnt
                Range("A1").Select
                wb_spread.Activate
                ws_spread0.Select
                
                Application.ScreenUpdating = True
                
                ' 集計表サマリーファイルを保存してクローズ
                spread_fn = Replace(summary_fn, "sum", "集計表")
                Open file_path & "\SUM\" & spread_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(spread_fn).Close
                End If
                If Dir(file_path & "\SUM\" & spread_fn) <> "" Then
                    Kill file_path & "\SUM\" & spread_fn
                End If
                
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=file_path & "\SUM\" & spread_fn
                ActiveWorkbook.Close
                Application.DisplayAlerts = True

                sum_cnt = sum_cnt + 1
                Application.Visible = True
                Unload Form_Progress
            Next n_cnt
            
            Set wb_spread = Nothing
            Set ws_spread0 = Nothing
            Set ws_spread1 = Nothing
            Set ws_spread2 = Nothing
            Set ws_spread3 = Nothing
            
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
              ws_mainmenu.Cells(41, 6) = "25"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 25"
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
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計表Excelファイルの作成：対象ファイル［SUMフォルダ内の" & sum_cnt - 1 & "個の集計サマリーデータ］"
            Close #1
            Call Finishing_Mcs2017
            MsgBox sum_cnt - 1 & "個の集計表Excelファイルが完成しました。", vbInformation, "MCS 2020 - Spreadsheet_Creation"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If
    
' 集計表Excelファイル１回作成処理
step00:
    wb.Activate
    ws_mainmenu.Select
    summary_fn = Application.GetOpenFilename("集計サマリーファイル,*.xlsx", , "集計サマリーファイルを開く")
    If summary_fn = "False" Then
        ' キャンセルボタンの処理
        Call Finishing_Mcs2017
        End
    ElseIf summary_fn = "" Then
        MsgBox "集計表Excelファイルを作成する［集計サマリーファイル］を選択してください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        GoTo step00
    ElseIf InStr(summary_fn, "_sum") = 0 Then
        MsgBox "集計表Excelファイルを作成する［集計サマリーファイル］を選択してください。", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        GoTo step00
    End If
    
    Open summary_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(summary_fn).Close
    Else
        Workbooks.Open summary_fn
    End If
    
    ' フルパスからフォルダ名の取得
    yen_pos = InStrRev(summary_fn, "\")
    summary_fd = Left(summary_fn, yen_pos - 1)
    
    ' フルパスからファイル名の取得
    summary_fn = Dir(summary_fn)
    
    Set wb_spread = Workbooks(summary_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

    wb_spread.Activate
    Sheets(Array("Ｎ％表", "Ｎ表", "％表")).Select
    Cells.Select
    
    ' 集計表シート全体のスタイル設定
    ActiveWindow.DisplayGridlines = False
    With Selection.Font
        .Name = Mid(mcs_ini(2), 8)      ' 日本語フォントをセット
        .Size = Mid(mcs_ini(3), 13)     ' 日本語フォントサイズをセット
    End With
    
    ' 集計表シート全体の表№のスタイル設定
    Columns(1).Select
    With Selection
        .ColumnWidth = 4.88
    End With
    
    ' 集計表シート全体の表側カテゴリー番号のスタイル設定
    Columns(3).Select
    With Selection
        .HorizontalAlignment = xlRight
        .ColumnWidth = 4.25
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With
    
    Range("A1").Select
    ws_spread0.Select
    ws_spread0.PageSetup.Orientation = xlLandscape
    hyo_cnt = ws_spread0.Cells(Rows.Count, setup_col).End(xlUp).Row - 1
    
    face_flag = 0       ' クロス集計表の有無判定用フラグ
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - 集計表Excelファイルの作成"
    Form_Progress.Repaint
    progress_msg = "集計表Excelファイルの作成をキャンセルしました。"
    Application.Visible = False
    AppActivate Form_Progress.Caption
    
    ' Ｎ％表の処理
    ws_spread1.Select
    ws_spread1.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP1/5 集計表Excelファイル（Ｎ％表）作成中" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        Call int_spreadsheet
    Next i_cnt
    
    ws_spread1.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "游ゴシック"
      .Size = 8
    End With
    ws_spread1.Columns(2).Hidden = True     ' MCODE列の非表示
    Range("A1").Select
    
    ' Ｎ表の処理
    ws_spread2.Select
    ws_spread2.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
'---------------------------------
    np_max_row = ActiveCell.Row
'---------------------------------
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP2/5 集計表Excelファイル（Ｎ表）作成中" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        Call ken_spreadsheet
    Next i_cnt
    
    ws_spread2.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "游ゴシック"
      .Size = 8
    End With
    ws_spread2.Columns(2).Hidden = True     ' MCODE列の非表示
    Range("A1").Select

    ' ％表の処理
    ws_spread3.Select
    ws_spread3.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
'---------------------------------
    np_max_row = ActiveCell.Row
'---------------------------------
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP3/5 集計表Excelファイル（％表）作成中" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        Call per_spreadsheet
    Next i_cnt
    
    ws_spread3.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "游ゴシック"
      .Size = 8
    End With
    ws_spread3.Columns(2).Hidden = True     ' MCODE列の非表示
    Range("A1").Select
    
    Form_Progress.Label1.Caption = "100%"
    DoEvents
    Form_Progress.Label2.Caption = "STEP4/5 集計表Excelファイル（目次）作成中..."
    Form_Progress.Label3.Caption = "[1/1ファイル]"
    waitTime = Now + TimeValue("0:00:01")
    Application.Wait waitTime
    
    ' 単純集計時の表側幅の調整
    If face_flag = 0 Then
        wb_spread.Activate
        Sheets(Array("Ｎ％表", "Ｎ表", "％表")).Select
        ws_spread1.Columns(5).ColumnWidth = 12.37
        ws_spread2.Columns(5).ColumnWidth = 12.37
        ws_spread3.Columns(5).ColumnWidth = 12.37
        Range("A1").Select
    End If
    
    ' 目次の処理
    wb_spread.Activate
    ws_spread0.Select
    ActiveWindow.DisplayGridlines = False
    Cells.Select
    With Selection.Font
        .Name = Mid(mcs_ini(2), 8)      ' 日本語フォントをセット
        .Size = 8                       ' フォントサイズをセット
    End With
    
    Columns(4).Select
    Columns(4).WrapText = True
    Columns(4).ColumnWidth = 42
    
    Columns(5).Select
    Columns(5).WrapText = True
    Columns(5).ColumnWidth = 42
    
    If WorksheetFunction.CountA(ws_spread0.Columns(6)) > 1 Then
        Columns(6).Select
        Columns(6).WrapText = True
        Columns(6).ColumnWidth = 42
    End If
    ws_spread0.Columns(3).Hidden = True     ' MCODE列の非表示
    
    ' 目次の罫線処理
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.LineStyle = xlContinuous
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    
    With Range(Selection, ActiveCell.SpecialCells(xlLastCell))
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
    With Range("A1:I1")
        .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    Range("G1:I1").Merge
    Range("G1").HorizontalAlignment = xlLeft
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select
    
    ' ウエイトバック集計時の説明文
    If ws_spread0.Cells(1, 1) = "連番ウあ" Then
        ws_spread0.Cells(1, 1) = "連番"
        Range("A1").End(xlDown).Select
        wgt_row = ActiveCell.Row
        ws_spread0.Cells(wgt_row + 1, 1) = "※ウエイトバック集計を行っているため、計算過程で小数点が生じますが、本集計表上の数値は四捨五入して整数表記しています。"
        Range("A1").Select
    End If
    
    ' 各シートの色を設定
    Sheets("Ｎ％表").Select
    With ActiveWorkbook.Sheets("Ｎ％表").Tab
        .Color = 10066431
        .TintAndShade = 0
    End With
    Sheets("Ｎ表").Select
    With ActiveWorkbook.Sheets("Ｎ表").Tab
        .Color = 10092441
        .TintAndShade = 0
    End With
    Sheets("％表").Select
    With ActiveWorkbook.Sheets("％表").Tab
        .Color = 16764057
        .TintAndShade = 0
    End With
    Sheets("目次").Select
    
' TOP1・GTソート処理 - 2018.10.04 追加
    ws_spread1.Select
    ws_spread1.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP5/5 集計表Excelファイル 最終調整中" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        Call top1_sort
    Next i_cnt
    Range("A1").Select
    wb_spread.Activate
    ws_spread0.Select
    
    Application.ScreenUpdating = True
    
    ' 集計サマリーファイルを保存してクローズ
    spread_fn = Replace(summary_fn, "sum", "集計表")
    Open file_path & "\SUM\" & spread_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(spread_fn).Close
    End If
    If Dir(file_path & "\SUM\" & spread_fn) <> "" Then
        Kill file_path & "\SUM\" & spread_fn
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=summary_fd & "\" & spread_fn
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Application.Visible = True
    Unload Form_Progress
    
    Set wb_spread = Nothing
    Set ws_spread0 = Nothing
    Set ws_spread1 = Nothing
    Set ws_spread2 = Nothing
    Set ws_spread3 = Nothing

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
      ws_mainmenu.Cells(41, 6) = "25"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 25"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計表Excelファイルの作成：対象ファイル［" & summary_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "集計表Excelファイルが完成しました。", vbInformation, "MCS 2020 - Spreadsheet_Creation"
End Sub

Private Sub int_spreadsheet()
' 件数＋％表のスタイル設定
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread1.Select
    
    ' 集計表１表あたりの開始行取得
    bgn_row = ActiveCell.Row

    ' 集計表１表あたりの開始列取得
    bgn_col = ActiveCell.Column
    
    ' 集計表１表あたりの最終行取得
    Selection.End(xlDown).Select
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' 件数欄の行列取得
    For k_cnt = bgn_row To fin_row
        If ws_spread1.Cells(k_cnt, 6) = "件数" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread1.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column
    
    Selection.End(xlToRight).Select

    ' 集計表１表あたりの最終列取得
    fin_col = ActiveCell.Column

    ' 表題のスタイル設定
    ws_spread1.Select
    With ws_spread1.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '［設問形式］と［構成比母数］のスタイル設定 - 2018.05.24 追加
    Range(ws_spread1.Cells(ken_row, 4), ws_spread1.Cells(ken_row, 5)).Merge
    With ws_spread1.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' 表頭カテゴリー番号のスタイル設定
    With ws_spread1.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' 表頭項目のスタイル設定
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' 全体行のスタイル設定
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "　全　体" Then
        total_flag = 1
        ws_spread1.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread1.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' 全体行の行列取得
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread1.Cells(zen_row, zen_col), ws_spread1.Cells(zen_row + 1, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread1.Cells(zen_row, zen_col + 1), ws_spread1.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' 集計値のスタイル設定
    With Range(ws_spread1.Cells(ken_row + 1, ken_col), ws_spread1.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' 英数字フォントをセット
        .Font.Size = Mid(mcs_ini(5), 13)     ' 英数字フォントサイズをセット
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' 表側表題のスタイル設定
    If ws_spread1.Cells(ken_row + 3, ken_col - 3) <> "" Then
        face_flag = 1
        cross_flag = 1
        
        ' 全体欄の有無によって、座標を調整
        If total_flag = 1 Then
            adjust_row = 3
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread1.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread1.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread1.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread1.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' 集計表１表あたりの罫線処理
    With Range(ws_spread1.Cells(ken_row, ken_col - 2), ws_spread1.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread1.Cells(ken_row, ken_col - 2), ws_spread1.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread1.Cells(ken_row, ken_col), ws_spread1.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' 表側項目の罫線処理
    If cross_flag = 1 Then
        For f_cnt = ken_row + 3 To fin_row - 1 Step 2
            With Range(ws_spread1.Cells(f_cnt, ken_col - 1), ws_spread1.Cells(f_cnt + 1, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' 次の集計表の表№に移動
    ws_spread1.Cells(fin_row + 2, 1).Select

End Sub

Private Sub ken_spreadsheet()
' 件数表のスタイル設定
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer

    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread2.Select
    
    ' 集計表１表あたりの開始行取得
    bgn_row = ActiveCell.Row

    ' 集計表１表あたりの開始列取得
    bgn_col = ActiveCell.Column

    Selection.End(xlDown).Select

    ' 集計表１表あたりの最終行取得
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' 件数欄の行列取得
    For k_cnt = bgn_row To fin_row
        If ws_spread2.Cells(k_cnt, 6) = "件数" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread2.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    Selection.End(xlToRight).Select

    ' 集計表１表あたりの最終列取得
    fin_col = ActiveCell.Column

    ' 表題のスタイル設定
    With ws_spread2.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '［設問形式］と［構成比母数］のスタイル設定 - 2018.05.24 追加
    Range(ws_spread2.Cells(ken_row, 4), ws_spread2.Cells(ken_row, 5)).Merge
    With ws_spread2.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' 表頭カテゴリー番号のスタイル設定
    With ws_spread2.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' 表頭項目のスタイル設定
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' 全体行のスタイル設定
    If ws_spread2.Cells(ken_row + 1, ken_col - 2) = "　全　体" Then
        total_flag = 1
        ws_spread2.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread2.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' 全体行の行列取得
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread2.Cells(zen_row, zen_col), ws_spread2.Cells(zen_row, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread2.Cells(zen_row, zen_col + 1), ws_spread2.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' 集計値のスタイル設定
    With Range(ws_spread2.Cells(ken_row + 1, ken_col), ws_spread2.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' 英数字フォントをセット
        .Font.Size = Mid(mcs_ini(5), 13)     ' 英数字フォントサイズをセット
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' 表側表題のスタイル設定
    If ws_spread2.Cells(ken_row + 2, ken_col - 3) <> "" Then
        cross_flag = 1
        
        ' 全体欄の有無によって、座標を調整
        If total_flag = 1 Then
            adjust_row = 2
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread2.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread2.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread2.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread2.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' 集計表１表あたりの罫線処理
    With Range(ws_spread2.Cells(ken_row, ken_col - 2), ws_spread2.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread2.Cells(ken_row, ken_col - 2), ws_spread2.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread2.Cells(ken_row, ken_col), ws_spread2.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' 表側項目の罫線処理
    If cross_flag = 1 Then
        For f_cnt = ken_row + 2 To fin_row Step 1
            With Range(ws_spread2.Cells(f_cnt, ken_col - 1), ws_spread2.Cells(f_cnt, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' 次の集計表の表№に移動
    ws_spread2.Cells(fin_row + 2, 1).Select

End Sub

Private Sub per_spreadsheet()
' ％表のスタイル設定
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread3.Select
    
    ' 集計表１表あたりの開始行取得
    bgn_row = ActiveCell.Row

    ' 集計表１表あたりの開始列取得
    bgn_col = ActiveCell.Column

    Selection.End(xlDown).Select

    ' 集計表１表あたりの最終行取得
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' 件数欄の行列取得
    For k_cnt = bgn_row To fin_row
        If ws_spread3.Cells(k_cnt, 6) = "件数" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread3.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    Selection.End(xlToRight).Select

    ' 集計表１表あたりの最終列取得
    fin_col = ActiveCell.Column

    ' 表題のスタイル設定
    With ws_spread3.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '［設問形式］と［構成比母数］のスタイル設定 - 2018.05.24 追加
    Range(ws_spread3.Cells(ken_row, 4), ws_spread3.Cells(ken_row, 5)).Merge
    With ws_spread3.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' 表頭カテゴリー番号のスタイル設定
    With ws_spread3.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' 表頭項目のスタイル設定
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' 全体行のスタイル設定
    If ws_spread3.Cells(ken_row + 1, ken_col - 2) = "　全　体" Then
        total_flag = 1
        ws_spread3.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread3.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' 全体行の行列取得
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread3.Cells(zen_row, zen_col), ws_spread3.Cells(zen_row, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread3.Cells(zen_row, zen_col + 1), ws_spread3.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' 集計値のスタイル設定
    With Range(ws_spread3.Cells(ken_row + 1, ken_col), ws_spread3.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' 英数字フォントをセット
        .Font.Size = Mid(mcs_ini(5), 13)     ' 英数字フォントサイズをセット
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' 表側表題のスタイル設定
    If ws_spread3.Cells(ken_row + 2, ken_col - 3) <> "" Then
        cross_flag = 1
        
        ' 全体欄の有無によって、座標を調整
        If total_flag = 1 Then
            adjust_row = 2
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread3.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread3.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread3.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread3.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' 集計表１表あたりの罫線処理
    With Range(ws_spread3.Cells(ken_row, ken_col - 2), ws_spread3.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread3.Cells(ken_row, ken_col - 2), ws_spread3.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread3.Cells(ken_row, ken_col), ws_spread3.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' 表側項目の罫線処理
    If cross_flag = 1 Then
        For f_cnt = ken_row + 2 To fin_row Step 1
            With Range(ws_spread3.Cells(f_cnt, ken_col - 1), ws_spread3.Cells(f_cnt, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' 次の集計表の表№に移動
    ws_spread3.Cells(fin_row + 2, 1).Select

End Sub

Private Sub top1_sort()
' TOP1・GTソート処理 - 2018.11.13 最終更新
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
'---------------------------------
    Dim val_range As Range
    Dim c_cnt As Long
    Dim cx_cnt As Long
    Dim face_cnt As Long
    Dim ct_row As Long
    Dim ct_col As Long
    Dim ct_cnt As Long
    Dim top1_ct As Double
    Dim top1_col As Long
    Dim ex_ct As Long
    Dim h_num As String
    Dim np_bgn_row As Long
    Dim np_bgn_col As Long
    Dim np_fin_row As Long
    Dim np_fin_col As Long
    Dim np_ken_row As Long
    Dim np_ken_col As Long
    Dim np_zen_row As Long
    Dim np_zen_col As Long
    Dim np_ct_row As Long
    Dim np_ct_col As Long
'---------------------------------
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread1.Select
    
    ' 集計表１表あたりの開始行取得
    bgn_row = ActiveCell.Row

    ' 集計表１表あたりの開始列取得
    bgn_col = ActiveCell.Column
    
    ' 集計表１表あたりの最終行取得
    Selection.End(xlDown).Select
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' 件数欄の行列取得
    For k_cnt = bgn_row To fin_row
        If ws_spread1.Cells(k_cnt, 6) = "件数" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread1.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    ' カテゴリー番号の行列取得
    ct_row = ken_row - 1
    ct_col = ken_col + 1
    
    Selection.End(xlToRight).Select

    ' 集計表１表あたりの最終列取得
    fin_col = ActiveCell.Column

    ' 集計表１表あたりの初期設定
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "　全　体" Then
        ' 全体行の行列取得
        ws_spread1.Cells(ken_row + 1, ken_col - 2).Select
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        
        ' Ｎ表・％表の座標取得
        h_num = ws_spread1.Cells(bgn_row, bgn_col)
        ws_spread2.Select
        Columns("A:A").Select
        Selection.Find(What:=h_num, after:=ActiveCell, LookIn:=xlFormulas, _
         lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
         MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        np_bgn_row = ActiveCell.Row
        np_bgn_col = ActiveCell.Column
    
        ' Ｎ表・％表の集計表１表あたりの最終行取得
        Selection.End(xlDown).Select
        If ActiveCell.Row = 1048576 Then
            np_fin_row = np_max_row
        Else
            np_fin_row = ActiveCell.Row - 2
        End If
        Selection.End(xlUp).Select
    
        ' Ｎ表・％表の件数欄の行列取得
        For k_cnt = np_bgn_row To np_fin_row
            If ws_spread2.Cells(k_cnt, 6) = "件数" Then
                np_ken_row = k_cnt
                Exit For
            End If
        Next k_cnt
        ws_spread2.Cells(np_ken_row, 6).Select
        np_ken_col = ActiveCell.Column

        ' Ｎ表・％表のカテゴリー番号の行列取得
        np_ct_row = np_ken_row - 1
        np_ct_col = np_ken_col + 1
        Range("A1").Select
    End If

    ' GTソートの処理 - 2018.10.10
    If Mid(ws_spread1.Cells(bgn_row + 2, bgn_col + 1), 1, 1) = "Y" Then
        ws_spread1.Select
        ct_cnt = 0
        For c_cnt = ct_col To fin_col    ' カテゴリー数をカウント
            If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                Exit For
            ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                Exit For
            End If
            ct_cnt = ct_cnt + 1
        Next c_cnt

        ' 除外CTの確認
        ex_ct = Val(Mid(ws_spread1.Cells(bgn_row + 2, bgn_col + 1), 2))

        ' 除外CTがCT数を超える場合はソート処理しない
        If ct_cnt >= ex_ct Then
            If ex_ct <> 0 Then
                ct_cnt = ex_ct - 1
            End If

            ' ソート処理
            Range(ws_spread1.Cells(ken_row - 1, ken_col + 1), ws_spread1.Cells(fin_row, ken_col + ct_cnt)).Select
            ws_spread1.Sort.SortFields.Clear
            ws_spread1.Sort.SortFields.Add Key:=Rows(zen_row), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            With ws_spread1.Sort
                .SetRange Range(ws_spread1.Cells(ken_row - 1, ken_col + 1), ws_spread1.Cells(fin_row, ken_col + ct_cnt))
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlLeftToRight
                .SortMethod = xlStroke
                .Apply
            End With
            
            ' Ｎ％表からＮ表・％表へ振り分け
            For c_cnt = 0 To ct_cnt
                ws_spread2.Cells(np_ken_row - 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row - 1, ken_col + c_cnt)
                ws_spread3.Cells(np_ken_row - 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row - 1, ken_col + c_cnt)
                ws_spread2.Cells(np_ken_row, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row, ken_col + c_cnt)
                ws_spread3.Cells(np_ken_row, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row, ken_col + c_cnt)
                If c_cnt = 0 Then
                    ws_spread2.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                    ws_spread3.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                Else
                    ws_spread2.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                    ws_spread3.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                End If
            Next c_cnt
        
            ' クロス集計表の縦展開処理
            If face_flag = 1 Then
                ' Ｎ％表からＮ表へ振り分け
                face_cnt = 1
                For cx_cnt = zen_row + 2 To fin_row - 1 Step 2
                    For c_cnt = 0 To ct_cnt
                        If c_cnt = 0 Then
                            ws_spread2.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        Else
                            ws_spread2.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        End If
                    Next c_cnt
                    face_cnt = face_cnt + 1
                Next cx_cnt
                ' Ｎ％表から％表へ振り分け
                face_cnt = 1
                For cx_cnt = zen_row + 2 To fin_row - 1 Step 2
                    For c_cnt = 0 To ct_cnt
                        If c_cnt = 0 Then
                            ws_spread3.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        Else
                            ws_spread3.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        End If
                    Next c_cnt
                    face_cnt = face_cnt + 1
                Next cx_cnt
            End If
        End If
    End If

    ' 全体行のスタイル設定
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "　全　体" Then
        ' TOP1（全体）の処理 - 2018.09.28 追加
        ws_spread1.Select
        If ws_spread1.Cells(bgn_row + 1, bgn_col + 1) <> "" Then
            ct_cnt = 0
            For c_cnt = ct_col To fin_col    ' カテゴリー数をカウント
                If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                    Exit For
                ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                    Exit For
                End If
                ct_cnt = ct_cnt + 1
            Next c_cnt
            ws_spread1.Select
            If ct_cnt <> 0 Then    ' ct_cntが［0］なら、対象なし（カテゴリーがない設問…ありえないと思うけど…）
                Set val_range = Range(ws_spread1.Cells(ct_row + 2, ct_col), ws_spread1.Cells(ct_row + 2, ct_col + ct_cnt - 1))
                top1_ct = Application.WorksheetFunction.Max(val_range)
                Set val_range = Nothing
                For top1_col = ct_col To (ct_col + ct_cnt - 1)
                    If ws_spread1.Cells(ct_row + 2, top1_col) = top1_ct Then
                        If ws_spread1.Cells(ct_row + 2, top1_col).Value <> 0 Then    '件数が ［0件］なら着色しない
                            Range(ws_spread1.Cells(ct_row + 2, top1_col), ws_spread1.Cells(ct_row + 3, top1_col)).Interior.Color = 8420607
                            Range(ws_spread2.Cells(np_ct_row + 2, top1_col), ws_spread2.Cells(np_ct_row + 2, top1_col)).Interior.Color = 8420607
                            Range(ws_spread3.Cells(np_ct_row + 2, top1_col), ws_spread3.Cells(np_ct_row + 2, top1_col)).Interior.Color = 8420607
                        End If
                    End If
                Next top1_col
            End If
        End If
    End If
    
    ' 表側表題のスタイル設定
    If ws_spread1.Cells(ken_row + 3, ken_col - 3) <> "" Then
        ' TOP1（表側項目）の処理 - 2018.10.03 追加
        ct_cnt = 0
        For c_cnt = ct_col To fin_col    ' カテゴリー数をカウント
            If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                Exit For
            ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                Exit For
            End If
            ct_cnt = ct_cnt + 1
        Next c_cnt
        If ct_cnt <> 0 Then    ' ct_cntが［0］なら、対象なし（カテゴリーがない設問…ありえないと思うけど…）
            If ws_spread1.Cells(bgn_row + 1, bgn_col + 1) = "A" Then
                If face_cnt = 0 Then
                    face_cnt = 1
                    For f_cnt = zen_row + 3 To fin_row Step 2
                        face_cnt = face_cnt + 1
                    Next f_cnt
                End If
                For f_cnt = 1 To face_cnt - 1
                    If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), zen_col - 1) <> "N/A" Then
                        Set val_range = Range(ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), ct_col), ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), ct_col + ct_cnt - 1))
                        top1_ct = Application.WorksheetFunction.Max(val_range)
                        Set val_range = Nothing
                        For top1_col = ct_col To (ct_col + ct_cnt - 1)
                            If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col) = top1_ct Then
                                If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col).Value <> 0 Then    '件数が ［0件］なら着色しない
                                    Range(ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col), ws_spread1.Cells(ct_row + 3 + (f_cnt * 2), top1_col)).Interior.Color = 8420607
                                    Range(ws_spread2.Cells(np_ct_row + 2 + f_cnt, top1_col), ws_spread2.Cells(np_ct_row + 2 + f_cnt, top1_col)).Interior.Color = 8420607
                                    Range(ws_spread3.Cells(np_ct_row + 2 + f_cnt, top1_col), ws_spread3.Cells(np_ct_row + 2 + f_cnt, top1_col)).Interior.Color = 8420607
                                End If
                            End If
                        Next top1_col
                    End If
                Next f_cnt
            End If
        End If
    End If

    ' 次の集計表の表№に移動
    ws_spread1.Cells(fin_row + 2, 1).Select

End Sub

