Attribute VB_Name = "Module56"
Option Explicit
    Dim spread_fn As String
    Dim spread_fd As String
    
    Dim csv_fn As String
    Dim csv_fd As String
    
    Dim wb_spread As Workbook
    Dim ws_spread0 As Worksheet
    Dim ws_spread1 As Worksheet
    Dim ws_spread2 As Worksheet
    Dim ws_spread3 As Worksheet


Sub Legacycsv_spreadsheet()
    Dim rc As Integer
    Dim yen_pos As Long
    Dim i_cnt As Long, s_cnt As Long
    Dim max_row As Long, max_col As Long
'2018/06/01 - 追記 ==========================
    Dim r_code As Integer
    Dim spd_tab() As String
    Dim spd_file As String
    Dim spd_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long
'--------------------------------------------------------------------------------------------------'
'　レガシー版集計表CSVファイルの作成 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.11.15　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "レガシー版集計表CSVファイルの作成中..."
    
    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    ' SUMフォルダ内の*_集計表.xlsx形式のファイル数をカウント
    spd_cnt = 0
    spd_file = Dir(file_path & "\SUM\*_集計表.xlsx")
    Do Until spd_file = ""
        DoEvents
        spd_cnt = spd_cnt + 1
        spd_file = Dir()
    Loop
    
    ' SUMフォルダ内の*_集計表.xlsx形式のファイル名を配列にセット
    ReDim spd_tab(spd_cnt)
    spd_file = Dir(file_path & "\SUM\*_集計表.xlsx")
    For fn_cnt = 1 To spd_cnt
        DoEvents
        spd_tab(fn_cnt) = spd_file
        spd_file = Dir()
    Next fn_cnt
    fn_cnt = spd_cnt
  
    rc = MsgBox("集計表Excelファイルから、レガシー版集計表CSVファイルを作成します。" & vbCrLf & "作成対象となる集計表Excelファイルはありますか。" _
      & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "集計表CSVファイルを作成するために必要な集計表Excelファイルがない場合は「いいえ」を選択してください。", vbYesNoCancel + vbQuestion, "集計表Excelファイル作成の確認")
    If rc = vbNo Then
        MsgBox "集計表Excelファイルを作成します。集計サマリーデータを選択してください。"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        Call Finishing_Mcs2017
        End
    End If

' レガシー版集計表CSVファイル複数作成処理（未着手、まずは１回処理から着手しています）
    If spd_cnt > 0 Then
        r_code = MsgBox("SUMフォルダ内にある" & fn_cnt & "個の集計表Excelファイルから、" & vbCrLf & "一括してレガシー版集計表CSVファイルを作成しますか。" _
         & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "SUMフォルダ内の［*_集計表.xlsx形式］のファイル数を" & vbCrLf & "表示しています。" _
         & vbCrLf & "「はい」　→ 集計表Excelファイルを一括処理" & vbCrLf & "「いいえ」→ 集計表Excelファイルを選択してから処理", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Csv_Spreadsheet")
        If r_code = vbYes Then
            spd_cnt = 1
            For n_cnt = 1 To fn_cnt
                DoEvents
                wb.Activate
                ws_mainmenu.Select
                spread_fn = spd_tab(n_cnt)
                
                Open file_path & "\SUM\" & spread_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(spread_fn).Close
                Else
                    Workbooks.Open spread_fn
                End If

                ' ファイル名から拡張子以外の取得
                csv_fn = Left(spread_fn, InStr(spread_fn, "_集計表") - 1)

                Set wb_spread = Workbooks(spread_fn)
                Set ws_spread0 = wb_spread.Worksheets(1)
                Set ws_spread1 = wb_spread.Worksheets(2)
                Set ws_spread2 = wb_spread.Worksheets(3)
                Set ws_spread3 = wb_spread.Worksheets(4)

                Application.DisplayAlerts = False
                wb_spread.Activate

                csv_fd = file_path & "\SUM\CSV\"
                If Dir(csv_fd, vbDirectory) = "" Then
                    MkDir csv_fd
                End If

                Sheets("目次").Select
                Call cells_format
                Call index_format
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_目次.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("Ｎ％表").Select
                Call cells_format
                Call legacy_format1
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP表.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("Ｎ表").Select
                Call cells_format
                Call legacy_format2
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N表.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("％表").Select
                Call cells_format
                Call legacy_format3
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_P表.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False

                ActiveWorkbook.Close
                Application.DisplayAlerts = True

                spd_cnt = spd_cnt + 1
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
            ActiveSheet.Unprotect Password:=""
            ws_mainmenu.Cells(initial_row, initial_col).Locked = False
            If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
              ws_mainmenu.Cells(41, 6) = "28"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 28"
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
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計表CSVファイルの作成：対象ファイル［SUMフォルダ内の" & spd_cnt - 1 & "個の集計表Excelファイル］"
            Close #1
            Call Finishing_Mcs2017
            MsgBox spd_cnt - 1 & "個の集計表CSVファイルが完成しました。", vbInformation, "MCS 2020 - Csv_spreadsheet"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If

' レガシー版集計表CSVファイル１回作成処理（１回処理から着手、鋭意作成中）
step00:
    wb.Activate
    ws_mainmenu.Select
' うるさいので、下記メッセージをコメントアウト
'    MsgBox "レガシー版集計表CSVファイルを作成する集計表Excelファイルを選択してください。"
    spread_fn = Application.GetOpenFilename("集計表Excelファイル,*.xlsx", , "集計表Excelファイルを開く")
    If spread_fn = "False" Then
        ' キャンセルボタンの処理
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf spread_fn = "" Then
        MsgBox "集計表Excelファイルを選択してください。", vbExclamation, "MCS 2020 - Print_spreadsheet"
        GoTo step00
    ElseIf InStr(spread_fn, "_集計表") = 0 Then
        MsgBox "集計表Excelファイルを選択してください。", vbExclamation, "MCS 2020 - Print_spreadsheet"
        GoTo step00
    End If

    Open spread_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(spread_fn).Close
    Else
        Workbooks.Open spread_fn
    End If
    
    ' フルパスからフォルダ名の取得
    yen_pos = InStrRev(spread_fn, "\")
    spread_fd = Left(spread_fn, yen_pos - 1)
    
    ' フルパスからファイル名の取得
    spread_fn = Dir(spread_fn)
    
    ' ファイル名から拡張子以外の取得
    csv_fn = Left(spread_fn, InStr(spread_fn, "_集計表") - 1)
    
    Set wb_spread = Workbooks(spread_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

' レガシー版集計表CSVファイル作成ここから
    
    Application.DisplayAlerts = False
    wb_spread.Activate

    csv_fd = spread_fd & "\CSV\"
    If Dir(csv_fd, vbDirectory) = "" Then
        MkDir csv_fd
    End If
    
    Sheets("目次").Select
    Call cells_format
    Call index_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_目次.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("Ｎ％表").Select
    Call cells_format
    Call legacy_format1
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP表.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("Ｎ表").Select
    Call cells_format
    Call legacy_format2
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N表.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("％表").Select
    Call cells_format
    Call legacy_format3
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_P表.csv", _
        FileFormat:=xlCSV, CreateBackup:=False

    ActiveWorkbook.Close
    Application.DisplayAlerts = True

' 集計表CSVファイル作成ここまで
    
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
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "28"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 28"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計表CSVファイルの作成：対象ファイル［" & spread_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "レガシー版集計表CSVファイルが完成しました。", vbInformation, "MCS 2020 - Csv_spreadsheet"
End Sub

Private Sub cells_format()
    Cells.Select
    With Selection
        .ClearFormats
    End With
    Range("A1").Select
End Sub

Private Sub index_format()
' 目次フォーマット調整
    Cells.Replace What:="" & Chr(10) & "", Replacement:="＆", lookat:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Private Sub legacy_format1()
' Ｎ％表フォーマット調整
    Dim bgn_row As Long, bgn_col As Long
    Dim fin_row As Long, fin_col As Long
    Dim max_row As Long, max_col As Long
    Dim ken_row As Long, ken_col As Long
    Dim k_cnt As Long
    Dim r_cnt As Long
    Dim p_cnt As Long
    Dim wk_row As Long
    Dim aj_row As Long
    Dim del_row As Long
    
    Dim sel_cm(6) As String
    Dim s_cnt As Long
    Dim findNo As Long
    Dim sel_flag As Integer
    
    Dim f_cnt As Long
    
    ' B列の削除
    Columns("B").Delete
    
    ' シートの最終行取得
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' 集計表１表あたりの開始行取得
            bgn_row = ActiveCell.Row
            ' 集計表１表あたりの開始列取得
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            'セレクトの有無確認とセレクトコメントの取得
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "【" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "】")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' 集計表１表あたりの最終行取得
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' 件数欄の行列取得
            For k_cnt = bgn_row To fin_row
                If ws_spread1.Cells(k_cnt, 5) = "件数" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' 集計表１表あたりの最終列取得
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' 表題コメントのセット
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "表題"
            Cells(ken_row + 1, ken_col - 1) = "合計"
            Cells(ken_row, ken_col - 3).Select
            
            ' セレクトコメントのセット
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "集計条件"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "：", "…")
                Next p_cnt
            End If
            
            ' 表側項目番号の処理
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' 表頭の「合計」「標準偏差」の処理
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "標準偏差" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "合計" Then
                    Cells(ken_row, p_cnt) = "実数合計"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' 件数と統計量（実数項目）の調整 - 2019.12.12
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "件数" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "平均" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "最小値" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "第１四分位" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "中央値" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "第３四分位" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "最大値" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "最頻値" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "標準偏差" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "実数合計" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
            Next p_cnt
            
            ' 集計表番号の処理と不要な行の削除処理
            Cells(bgn_row, bgn_col).Select
            Cells(bgn_row, bgn_col) = Cells(bgn_row, bgn_col) & "  -0-0000"
            Selection.Copy
            Range(Cells(ken_row, 1), Cells(fin_row + (s_cnt - 1), 1)).Select
            ActiveSheet.Paste
            Range(Rows(bgn_row), Rows(bgn_row + (s_cnt - 1) + 1)).Select
            aj_row = Selection.Rows.Count
'            del_row = del_row + aj_row
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlUp
            
            ' 次の集計表の始点をセレクト
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C列の削除
    Columns("C").Delete

End Sub

Private Sub legacy_format2()
' Ｎ表フォーマット調整
    Dim bgn_row As Long, bgn_col As Long
    Dim fin_row As Long, fin_col As Long
    Dim max_row As Long, max_col As Long
    Dim ken_row As Long, ken_col As Long
    Dim k_cnt As Long
    Dim r_cnt As Long
    Dim p_cnt As Long
    Dim wk_row As Long
    Dim aj_row As Long
    Dim del_row As Long
    
    Dim sel_cm(6) As String
    Dim s_cnt As Long
    Dim findNo As Long
    Dim sel_flag As Integer
    
    ' B列の削除
    Columns("B").Delete
    
    ' シートの最終行取得
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' 集計表１表あたりの開始行取得
            bgn_row = ActiveCell.Row
            ' 集計表１表あたりの開始列取得
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            'セレクトの有無確認とセレクトコメントの取得
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "【" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "】")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' 集計表１表あたりの最終行取得
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' 件数欄の行列取得
            For k_cnt = bgn_row To fin_row
                If ws_spread2.Cells(k_cnt, 5) = "件数" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' 集計表１表あたりの最終列取得
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' 表題コメントのセット
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "表題"
            Cells(ken_row + 1, ken_col - 1) = "合計"
            Cells(ken_row, ken_col - 3).Select
            
            ' セレクトコメントのセット
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "集計条件"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "：", "…")
                Next p_cnt
            End If
            
            ' 表側項目番号の処理
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' 表頭の「合計」「標準偏差」の処理
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "標準偏差" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "合計" Then
                    Cells(ken_row, p_cnt) = "実数合計"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' 集計表番号の処理と不要な行の削除処理
            Cells(bgn_row, bgn_col).Select
            Cells(bgn_row, bgn_col) = Cells(bgn_row, bgn_col) & "  -0-0000"
            Selection.Copy
            Range(Cells(ken_row, 1), Cells(fin_row + (s_cnt - 1), 1)).Select
            ActiveSheet.Paste
            Range(Rows(bgn_row), Rows(bgn_row + (s_cnt - 1) + 1)).Select
            aj_row = Selection.Rows.Count
'            del_row = del_row + aj_row
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlUp
            
            ' 次の集計表の始点をセレクト
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C列の削除
    Columns("C").Delete

End Sub

Private Sub legacy_format3()
' ％表フォーマット調整
    Dim bgn_row As Long, bgn_col As Long
    Dim fin_row As Long, fin_col As Long
    Dim max_row As Long, max_col As Long
    Dim ken_row As Long, ken_col As Long
    Dim k_cnt As Long
    Dim r_cnt As Long
    Dim p_cnt As Long
    Dim wk_row As Long
    Dim aj_row As Long
    Dim del_row As Long
    
    Dim sel_cm(6) As String
    Dim s_cnt As Long
    Dim findNo As Long
    Dim sel_flag As Integer
    
    ' B列の削除
    Columns("B").Delete
    
    ' シートの最終行取得
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' 集計表１表あたりの開始行取得
            bgn_row = ActiveCell.Row
            ' 集計表１表あたりの開始列取得
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            'セレクトの有無確認とセレクトコメントの取得
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "【" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "】")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' 集計表１表あたりの最終行取得
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' 件数欄の行列取得
            For k_cnt = bgn_row To fin_row
                If ws_spread3.Cells(k_cnt, 5) = "件数" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' 集計表１表あたりの最終列取得
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' 表題コメントのセット
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "表題"
            Cells(ken_row + 1, ken_col - 1) = "合計"
            Cells(ken_row, ken_col - 3).Select
            
            ' セレクトコメントのセット
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "集計条件"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "：", "…")
                Next p_cnt
            End If
            
            ' 表側項目番号の処理
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' 表頭の「合計」「標準偏差」の処理
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "標準偏差" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "合計" Then
                    Cells(ken_row, p_cnt) = "実数合計"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' 集計表番号の処理と不要な行の削除処理
            Cells(bgn_row, bgn_col).Select
            Cells(bgn_row, bgn_col) = Cells(bgn_row, bgn_col) & "  -0-0000"
            Selection.Copy
            Range(Cells(ken_row, 1), Cells(fin_row + (s_cnt - 1), 1)).Select
            ActiveSheet.Paste
            Range(Rows(bgn_row), Rows(bgn_row + (s_cnt - 1) + 1)).Select
            aj_row = Selection.Rows.Count
'            del_row = del_row + aj_row
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlUp
            
            ' 次の集計表の始点をセレクト
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C列の削除
    Columns("C").Delete

End Sub


