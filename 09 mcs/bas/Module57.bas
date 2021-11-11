Attribute VB_Name = "Module57"
Option Explicit
    Dim wb_tabinst As Workbook
    Dim ws_tabinst As Worksheet
    Dim tabinst_fn As String
    Dim face_qcode As String
    Dim period_pos As Long

    Dim t_seq As Long    '表№
    Dim t_r As Long      '表№行カウント
    Dim t_ra As Long     '実数指定行カウント
    Dim ct_f As Integer  '実数カテゴライズフラグ

    Dim t_crs As String  '第3軸のQCODE
    Dim t_cnt As Long    '第3軸のカテゴリー数

Sub Triplecross_Setting()
    Dim d_index As Long
    Dim r_code As Integer
    Dim i_cnt As Long
'--------------------------------------------------------------------------------------------------'
'　３重クロス用集計設定ファイルの作成　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2019.10.02　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check

    wb.Activate
    ws_mainmenu.Select
step00:
    tabinst_fn = InputBox("作成する集計設定ファイルのファイル名を" & vbCrLf & "入力してください。" & vbCrLf & vbCrLf & "【例】A01030C01.xlsx など", "MCS 2020 - 3重クロス用集計設定ファイルの作成")
    If tabinst_fn = "" Then
        Application.StatusBar = False
        End
    End If

    period_pos = InStrRev(tabinst_fn, ".")
    If period_pos > 0 Then
        If LCase(Mid(tabinst_fn, period_pos + 1)) <> "xlsx" Then
            MsgBox "ファイル形式が正しくありません。" _
             & vbCrLf & "拡張子は xlsx を指定してください。", vbExclamation, "MCS 2020 - Tabulation_Setting"
            GoTo step00
        End If
    Else
        tabinst_fn = tabinst_fn & ".xlsx"
    End If
    
    Call Setup_Hold
'2019.10.9 - 追加処理---------------------------------------------------
    ws_setup.Select
    face_qcode = ""
    face_qcode = InputBox("表側に設定するQCODEを入力してください。" & vbCrLf & vbCrLf & "【例】F01 など" & vbCrLf & "※未入力の場合は、単純集計の集計設定となります。", "MCS 2020 - 3重クロス用集計設定ファイルの作成")
    
    If StrPtr(face_qcode) = 0 Then  'キャンセル処理
        wb.Activate
        ws_mainmenu.Select
        Application.StatusBar = False
        End
    End If
    
    If face_qcode <> "" Then
        d_index = Qcode_Match(face_qcode)
        If (q_data(d_index).q_format = "R") Or (q_data(d_index).q_format = "H") _
         Or (q_data(d_index).q_format = "C") Or (q_data(d_index).q_format = "F") _
         Or (q_data(d_index).q_format = "O") Then
            face_qcode = ""
        End If
    End If
'-----------------------------------------------------------------------
    
step10:
    ws_setup.Select
    t_crs = ""
    t_crs = InputBox("第3軸（集計条件軸）に設定するQCODEを入力してください。" & vbCrLf & vbCrLf & "【例】KBN など", "MCS 2020 - 3重クロス用集計設定ファイルの作成")
    
    If StrPtr(t_crs) = 0 Then  'キャンセル処理
        wb.Activate
        ws_mainmenu.Select
        Application.StatusBar = False
        End
    End If
    
    If t_crs <> "" Then
        r_code = MsgBox("集計設定ファイルを第3軸のカテゴリーごとに分割して" & vbCrLf & "作成しますか？" _
         & vbCrLf & vbCrLf & "「はい」　→ 集計設定ファイルを分割作成" & vbCrLf & "「いいえ」→ 集計設定ファイルを1ファイルとして作成", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Spreadsheet_Creation")
        
        d_index = Qcode_Match(t_crs)
        If (q_data(d_index).q_format = "R") Or (q_data(d_index).q_format = "H") _
         Or (q_data(d_index).q_format = "C") Or (q_data(d_index).q_format = "F") _
         Or (q_data(d_index).q_format = "O") Then
            MsgBox "第3軸の形式を確認してください。" & vbCrLf & vbCrLf & _
             "【TIPS】" & vbCrLf & "第3軸に設定できるQCODEの形式は、" & vbCrLf & "［SA］［MA］［LMA］となっております。", vbInformation, "MCS 2020 - Triplecross_Setting"
            GoTo step10
        Else
            t_cnt = q_data(d_index).ct_count
        End If
    Else
        MsgBox "第3軸の指定がありません。" & vbCrLf & vbCrLf & _
         "【TIPS】" & vbCrLf & "通常の単純集計表・クロス集計表を作成する場合は、" & vbCrLf & "「集計設定ファイルの作成」から作成してください。", vbInformation, "MCS 2020 - Triplecross_Setting"
        GoTo step10
    End If
    
    Application.StatusBar = "3重クロス用集計設定ファイル 作成中..."
    Application.ScreenUpdating = False

    wb.Activate
    ws_mainmenu.Select

    If r_code = vbYes Then
        Call split_tabinst    ' 分割作成処理へ
        GoTo step99
    End If

' 集計設定ファイル1個作成処理
    Workbooks.Add
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & tabinst_fn
    Application.DisplayAlerts = True

    If Err.Number <> 0 Then
        ActiveWorkbook.Close
        Open file_path & "\3_FD\" & tabinst_fn For Append As #1
        Close #1
        If Err.Number = 70 Then
            MsgBox tabinst_fn & " は、すでに開かれています。" _
            & vbCrLf & "ファイルを閉じてから再実行してください。", vbExclamation, "MCS 2020 - Tabulation_Setting"
        End If
        End
    End If

    Set wb_tabinst = ActiveWorkbook
    Set ws_tabinst = wb_tabinst.ActiveSheet

    Call Inst_Header

' ここから集計設定ファイル作成のコーディング(´･ω･`)
    t_seq = 0
    For i_cnt = 1 To t_cnt
        For t_r = 3 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
            If (ws_setup.Cells(t_r, 1).Value <> "weight") And _
             (Mid(ws_setup.Cells(t_r, 1).Value, 1, 2) <> "SE") Then ' QCODE［weight］と［SE］はじまり（加工後セレクト）は集計設定ファイルに出力しない
                If Left(ws_setup.Cells(t_r, 1).Value, 1) <> "*" Then ' QCODE列の先頭アスタリスク行は処理しない
                    Select Case Left(ws_setup.Cells(t_r, 9).Value, 1)
                        Case "S", "M", "L"  ' SA,MA,LMAはQCODEを表頭に出力
                            If ws_setup.Cells(t_r, 2).Value = "" Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                            Else
                                For t_ra = t_r + 1 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                    If ws_setup.Cells(t_r, 2).Value = ws_setup.Cells(t_ra, 1).Value Then
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                                    End If
                                Next t_ra
                            End If
                            
                            ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                            ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                            
                            '2018/9/13 - 追加項目のための処理
                            d_index = Qcode_Match(ws_setup.Cells(t_r, 1).Value)
                            If Left(q_data(d_index).q_format, 1) = "S" Then
                                If q_data(d_index).ct_count <= 5 Then
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' 円グラフ
                                Else
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                End If
                            Else
                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                            End If
                        Case "R", "H"  ' RA,HCはカテゴライズ後のQCODEをループで探して表頭に出力
                            ct_f = 0
                            For t_ra = t_r To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value Then
                                    If Left(ws_setup.Cells(t_ra, 1).Value, 1) <> "*" Then
                                        ct_f = 1
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_ra, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_ra, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                                        
                                        ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                                        ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                                        
                                        ' 2018/9/13 - 追加項目のための処理
                                        d_index = Qcode_Match(ws_setup.Cells(t_ra, 1).Value)
                                        If Left(q_data(d_index).q_format, 1) = "S" Then
                                            If q_data(d_index).ct_count <= 5 Then
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' 円グラフ
                                            Else
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                            End If
                                        Else
                                            ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                        End If
                                    End If
                                End If
                            Next t_ra
                            If ct_f = 0 Then
                                For t_ra = 3 To t_r - 1
                                    If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value And t_r >= 3 Then
                                        ct_f = 1
                                    End If
                                Next t_ra
                            End If
                            If ct_f = 0 Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 1).Value
                                ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                            End If
                        Case "C", "F", "O"  ' CODE,FA,OAは集計設定ファイルに出力しない
                        Case Else
                    End Select
                End If
            End If
        Next t_r
    Next i_cnt

    wb_tabinst.Activate
    ws_tabinst.Select
    ws_tabinst.Range(Cells(7, 1), Cells(t_seq + 6, 25)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
    End With

    ' ファイル上部の処理と全体的な処理
    ws_tabinst.Cells(2, 1).Value = "集計するデータファイル名"

    Range("A2:C2").Select
    Selection.Merge

    Range("A2:G2").Select
    With Selection.Font
        .Color = 16724787
        .TintAndShade = 0
    End With
    
    ws_tabinst.Cells(2, 9).Value = "【ウエイト集計の設定】"
    ws_tabinst.Cells(2, 10).Value = "なし"
    ws_tabinst.Cells(2, 11).Value = "あり"

    Range("L2").Select
    ws_tabinst.Cells(2, 12).Value = "なし"
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$J$2:$K$2"
    End With
    
    Range("A2:C2").Select
    Selection.Copy
    Range("I2:K2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Range("E:Y").ColumnWidth = 7.13
    Range("E:Y").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Rows(1).RowHeight = 6.75
    Rows(2).RowHeight = 28.5
    Rows(3).RowHeight = 6.75
        
    Range("D2:G2").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws_tabinst.Cells(2, 4) = ws_mainmenu.Cells(3, 8) & "OT.xlsx"
    
    ActiveWindow.DisplayGridlines = False
    ws_tabinst.Cells(1, 1).Select

' 上書き保存してファイルを閉じる
    wb_tabinst.Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

    Set wb_tabinst = Nothing
    Set ws_tabinst = Nothing
    
' システムログの出力
step99:
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "22"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 22"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 集計設定ファイルの作成：作成ファイル［" & tabinst_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "集計設定ファイル " & tabinst_fn & " の作成が完了しました。", vbInformation, "MCS 2020 - Triplecross_Setting"
End Sub

Private Sub Inst_Header()
' ヘッダーの作成
    Cells.Select
    Selection.NumberFormatLocal = "@"
    With Selection.Font
        .Name = "Takaoゴシック"
        .Size = 11
    End With
    
    Range("A4:Y6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Range("A4:A6").Select
    Selection.Merge

    Range("B4:B5").Select
    Selection.Merge

    Range("C4:C5").Select
    Selection.Merge

    Range("D4:D5").Select
    Selection.Merge

    Range("E4:E5").Select
    Selection.Merge

    Range("F4:F5").Select
    Selection.Merge

    Range("G4:G5").Select
    Selection.Merge

    Range("A4:D6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    Range("E4:G6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 204
        .TintAndShade = 0
    End With

    Range("H4:P4").Select
    Selection.Merge
    Range("H4:P6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16772300
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -52429
        .TintAndShade = 0
    End With

    Range("Q4:R4").Select
    Selection.Merge
    Range("Q4:R6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434848
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 32768
        .TintAndShade = 0
    End With

    Range("S4:U4").Select
    Selection.Merge
    Range("S4:U6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 13158
        .TintAndShade = 0
    End With
    
    Range("V4:V6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10079487
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16764007
        .TintAndShade = 0
    End With

    Range("W4:X4").Select
    Selection.Merge
    Range("W4:X6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764006
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -6737152
        .TintAndShade = 0
    End With

    Range("Y4:Y4").Select
    Selection.Merge
    Range("Y4:Y6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777215
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

    Range("A4:Y6").Borders.LineStyle = xlContinuous
    Range("A6:Y6").Borders(xlEdgeTop).LineStyle = xlLineStyleNone

    Range("A4") = "表№"
    Range("B4") = "表側"
    Range("B6") = "(QCODE)"
    Range("C4") = "表頭"
    Range("C6") = "(QCODE)"
    Range("D4") = "実数"
    Range("D6") = "(QCODE)"
    Range("E4") = "表側表示"
    Range("E6") = "(N/E)"
    Range("F4") = "表頭NA"
    Range("F6") = "(N)"
    Range("G4") = "母数"
    Range("G6") = "(Y)"
    Range("H4") = "実数設問出力指示"
    Range("H5") = "合計"
    Range("H6") = "(Y)"
    Range("I5") = "平均"
    Range("I6") = "(Num/Y)"
    Range("J5") = "標準偏差"
    Range("J6") = "(Num/Y)"
    Range("K5") = "最小値"
    Range("K6") = "(Y)"
    Range("L5") = "第１四分位"
    Range("L6") = "(Y)"
    Range("M5") = "中央値"
    Range("M6") = "(Y)"
    Range("N5") = "第３四分位"
    Range("N6") = "(Y)"
    Range("O5") = "最大値"
    Range("O6") = "(Y)"
    Range("P5") = "最頻値"
    Range("P6") = "(Y)"
    Range("Q4") = "条件"
    Range("Q5") = "QCODE"
    Range("Q6") = "(QCODE)"
    Range("R5") = "値"
    Range("R6") = "(Num)"
    Range("S4") = "表示オプション"
    Range("S5") = "件数欄"
    Range("S6") = "(Y)"
    Range("T5") = "有効回答"
    Range("T6") = "(Y)"
    Range("U5") = "述べ回答"
    Range("U6") = "(Y)"

'2018/9/13 - 追加項目
    Range("V4") = "TOP1"
    Range("V5") = "マーキング"
    Range("V6") = "(Y/A)"
    Range("W4") = "GTソート"
    Range("W5") = "降順"
    Range("W6") = "(Y)"
    Range("X5") = "除外CT"
    Range("X6") = "(Num)"
    Range("Y4") = "グラフ"
    Range("Y5") = "種類"
    Range("Y6") = "(Num)"

    Range("A7").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub split_tabinst()
' 集計設定ファイル分割作成処理 - 2019.10.2 追記
    Dim fd As String
    Dim d_index As Long
    Dim r_code As Integer
    Dim i_cnt As Long
    
    For i_cnt = 1 To t_cnt
        Workbooks.Add
        Application.DisplayAlerts = False
        
        fd = Dir(file_path & "\3_FD\CRS", vbDirectory)
        If fd = "" Then
            MkDir file_path & "\3_FD\CRS"
        End If
        
        ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\CRS\" & Format(i_cnt, "00") & "_" & tabinst_fn
        Application.DisplayAlerts = True

        If Err.Number <> 0 Then
            ActiveWorkbook.Close
            Open file_path & "\3_FD\CRS\" & Format(i_cnt, "00") & "_" & tabinst_fn For Append As #1
            Close #1
            If Err.Number = 70 Then
                MsgBox Format(i_cnt, "00") & "_" & tabinst_fn & " は、すでに開かれています。" _
                 & vbCrLf & "ファイルを閉じてから再実行してください。", vbExclamation, "MCS 2020 - Tabulation_Setting"
            End If
            End
        End If
        
        Set wb_tabinst = ActiveWorkbook
        Set ws_tabinst = wb_tabinst.ActiveSheet

        Call Inst_Header

        t_seq = 0
        For t_r = 3 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
            If (ws_setup.Cells(t_r, 1).Value <> "weight") And _
             (Mid(ws_setup.Cells(t_r, 1).Value, 1, 2) <> "SE") Then ' QCODE［weight］と［SE］はじまり（加工後セレクト）は集計設定ファイルに出力しない
                If Left(ws_setup.Cells(t_r, 1).Value, 1) <> "*" Then ' QCODE列の先頭アスタリスク行は処理しない
                    Select Case Left(ws_setup.Cells(t_r, 9).Value, 1)
                        Case "S", "M", "L"  ' SA,MA,LMAはQCODEを表頭に出力
                            If ws_setup.Cells(t_r, 2).Value = "" Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                            Else
                                For t_ra = t_r + 1 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                    If ws_setup.Cells(t_r, 2).Value = ws_setup.Cells(t_ra, 1).Value Then
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                                    End If
                                Next t_ra
                            End If
                            
                            ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                            ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                            
                            '2018/9/13 - 追加項目のための処理
                            d_index = Qcode_Match(ws_setup.Cells(t_r, 1).Value)
                            If Left(q_data(d_index).q_format, 1) = "S" Then
                                If q_data(d_index).ct_count <= 5 Then
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' 円グラフ
                                Else
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                End If
                            Else
                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                            End If
                        Case "R", "H"  ' RA,HCはカテゴライズ後のQCODEをループで探して表頭に出力
                            ct_f = 0
                            For t_ra = t_r To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value Then
                                    If Left(ws_setup.Cells(t_ra, 1).Value, 1) <> "*" Then
                                        ct_f = 1
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_ra, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_ra, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                                        
                                        ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                                        ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                                        
                                        ' 2018/9/13 - 追加項目のための処理
                                        d_index = Qcode_Match(ws_setup.Cells(t_ra, 1).Value)
                                        If Left(q_data(d_index).q_format, 1) = "S" Then
                                            If q_data(d_index).ct_count <= 5 Then
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' 円グラフ
                                            Else
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                            End If
                                        Else
                                            ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' 横グラフ
                                        End If
                                    End If
                                End If
                            Next t_ra
                            If ct_f = 0 Then
                                For t_ra = 3 To t_r - 1
                                    If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value And t_r >= 3 Then
                                        ct_f = 1
                                    End If
                                Next t_ra
                            End If
                            If ct_f = 0 Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 1).Value
                                ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' 既定は小数点第１位で出力
                            End If
                        Case "C", "F", "O"  ' CODE,FA,OAは集計設定ファイルに出力しない
                        Case Else
                    End Select
                End If
            End If
        Next t_r
        wb_tabinst.Activate
        ws_tabinst.Select
        ws_tabinst.Range(Cells(7, 1), Cells(t_seq + 6, 25)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
        End With

        ' ファイル上部の処理と全体的な処理
        ws_tabinst.Cells(2, 1).Value = "集計するデータファイル名"

        Range("A2:C2").Select
        Selection.Merge

        Range("A2:G2").Select
        With Selection.Font
            .Color = 16724787
            .TintAndShade = 0
        End With

        ws_tabinst.Cells(2, 9).Value = "【ウエイト集計の設定】"
        ws_tabinst.Cells(2, 10).Value = "なし"
        ws_tabinst.Cells(2, 11).Value = "あり"

        Range("L2").Select
        ws_tabinst.Cells(2, 12).Value = "なし"
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$J$2:$K$2"
        End With
    
        Range("A2:C2").Select
        Selection.Copy
        Range("I2:K2").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        Range("E:Y").ColumnWidth = 7.13
        Range("E:Y").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        Rows(1).RowHeight = 6.75
        Rows(2).RowHeight = 28.5
        Rows(3).RowHeight = 6.75
        
        Range("D2:G2").Select
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        ws_tabinst.Cells(2, 4) = ws_mainmenu.Cells(3, 8) & "OT.xlsx"
    
        ActiveWindow.DisplayGridlines = False
        ws_tabinst.Cells(1, 1).Select
    
        ' 上書き保存してファイルを閉じる
        wb_tabinst.Activate
        Application.DisplayAlerts = False
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        Application.DisplayAlerts = True

        Set wb_tabinst = Nothing
        Set ws_tabinst = Nothing
    
    Next i_cnt
End Sub
