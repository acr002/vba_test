Attribute VB_Name = "Module55"
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


Sub Csv_spreadsheet()
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
'　集計表CSVファイルの作成 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.04.27　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "集計表CSVファイルの作成中..."
    
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
  
    rc = MsgBox("集計表Excelファイルから、集計表CSVファイルを作成します。" & vbCrLf & "作成対象となる集計表Excelファイルはありますか。" _
      & vbCrLf & vbCrLf & "【TIPS】" & vbCrLf & "集計表CSVファイルを作成するために必要な集計表Excelファイルがない場合は「いいえ」を選択してください。", vbYesNoCancel + vbQuestion, "集計表Excelファイル作成の確認")
    If rc = vbNo Then
        MsgBox "集計表Excelファイルを作成します。集計サマリーデータを選択してください。"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        Call Finishing_Mcs2017
        End
    End If

' 集計表CSVファイル複数作成処理
    If spd_cnt > 0 Then
        r_code = MsgBox("SUMフォルダ内にある" & fn_cnt & "個の集計表Excelファイルから、" & vbCrLf & "一括して集計表CSVファイルを作成しますか。" _
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
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_目次.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("Ｎ％表").Select
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP表.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("Ｎ表").Select
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N表.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("％表").Select
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
              ws_mainmenu.Cells(41, 6) = "27"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 27"
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

' 集計表CSVファイル１回作成処理
step00:
    wb.Activate
    ws_mainmenu.Select
    MsgBox "集計表CSVファイルを作成する集計表Excelファイルを選択してください。"
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

' 集計表CSVファイル作成ここから
    
    Application.DisplayAlerts = False
    wb_spread.Activate

    csv_fd = spread_fd & "\CSV\"
    If Dir(csv_fd, vbDirectory) = "" Then
        MkDir csv_fd
    End If
    
    Sheets("目次").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_目次.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("Ｎ％表").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP表.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("Ｎ表").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N表.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("％表").Select
    Call cells_format
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
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "27"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 27"
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
    MsgBox "集計表CSVファイルが完成しました。", vbInformation, "MCS 2020 - Csv_spreadsheet"
End Sub

Private Sub cells_format()
    Cells.Select
    With Selection
        .ClearFormats
    End With
    Range("A1").Select
End Sub
