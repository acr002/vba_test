Attribute VB_Name = "Module03"
Option Explicit

Sub Indata_Revision()
    Dim wb_revision As Workbook
    Dim ws_revision As Worksheet
    
    Dim wb_data As Workbook
    Dim ws_data As Worksheet
    
    Dim idata_fn As String
    Dim odata_fn As String
    Dim period_pos As Integer
    Dim max_row As Long, max_col As Long

    Dim rev_fn As String
    Dim rev_row As Long
    
    Dim gcode As String

    Dim dat_row As Long, dat_col As Long, rev_cnt As Long
    Dim FoundCell As Range
    Dim rev_sno As String
    Dim rev_qcode As String
    Dim rev_mact As Integer
    Dim rev_before As Variant
    Dim rev_after As Variant

    ' 修正指示のアドレス設定
    Const e_sno As Integer = 1      ' SampleNo
    Const e_qcode As Integer = 2    ' QCODE
    Const e_data As Integer = 5     ' 回答内容
    Const e_rst  As Integer = 6     ' 修正内容
'--------------------------------------------------------------------------------------------------'
'　入力データの修正　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　田中　義晃　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.05.15　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Setup_Hold
    Call Filepath_Get
    Call Setup_Check
    
    Application.ScreenUpdating = False

    wb.Activate
    ws_mainmenu.Select
    gcode = ws_mainmenu.Cells(gcode_row, gcode_col)

    ChDrive file_path & "\3_FD"
    ChDir file_path & "\3_FD"
    
step00:
    rev_fn = Application.GetOpenFilename("修正指示ファイル,*.xlsx", , "修正指示ファイルを開く")
    If rev_fn = "False" Then
        ' キャンセルボタンの処理
        Call Finishing_Mcs2017
        End
    ElseIf rev_fn = "" Then
        MsgBox "［修正指示ファイル］を選択してください。", vbExclamation, "MCS 2020 - Indata_Creation"
        GoTo step00
    ElseIf InStr(rev_fn, "_修正指示.xlsx") = 0 Then
        MsgBox "［修正指示ファイル］を選択してください。", vbExclamation, "MCS 2020 - Indata_Creation"
        GoTo step00
    End If
    
    Open rev_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(Dir(rev_fn)).Close
    Else
        Workbooks.Open rev_fn
    End If
    
    ' フルパスからファイル名の取得
    rev_fn = Dir(rev_fn)
    
    Set wb_revision = Workbooks(rev_fn)
    Set ws_revision = wb_revision.Worksheets(1)
    
    wb_revision.Activate
    ws_revision.Select
    rev_row = Cells(Rows.Count, 1).End(xlUp).Row
    Application.StatusBar = False
        
    ws_revision.Cells(5, 7).Value = "修正前"
    ws_revision.Cells(5, 8).Value = "修正後"

    idata_fn = ws_revision.Cells(2, 4)
    odata_fn = ws_revision.Cells(3, 4)
    
    ChDrive file_path & "\1_DATA"
    ChDir file_path & "\1_DATA"
    
    ' 修正ファイルの有無チェック
    If Dir(file_path & "\1_DATA\" & idata_fn) = "" Then
        MsgBox "修正ファイル名で設定されているファイル［" & idata_fn & "］が見つかりません。", vbExclamation, "MCS 2020 - Indata_Creation"
    Else
        Open idata_fn For Append As #1
        Close #1
        If Err.Number > 0 Then
            Workbooks(idata_fn).Close
        Else
            Workbooks.Open file_path & "\1_DATA\" & idata_fn
            Set wb_data = Workbooks(idata_fn)
            Set ws_data = wb_data.Worksheets(1)
            max_row = ws_data.Cells(Rows.Count, 1).End(xlUp).Row
            max_col = Cells(1, Columns.Count).End(xlToLeft).Column
        End If
    End If
    
    ' 出力ファイルの有無チェック
    If Dir(file_path & "\1_DATA\" & odata_fn) <> "" Then
        Kill file_path & "\1_DATA\" & odata_fn
    End If

    For rev_cnt = 6 To rev_row
        rev_sno = ws_revision.Cells(rev_cnt, e_sno).Value
        rev_qcode = ws_revision.Cells(rev_cnt, e_qcode).Value
        rev_mact = ws_revision.Cells(rev_cnt, e_qcode + 1).Value
        rev_before = ws_revision.Cells(rev_cnt, e_data).Value
        rev_after = ws_revision.Cells(rev_cnt, e_rst).Value

        wb_data.Activate
        ws_data.Select
        
        ' データのフォーマット確認
        If Cells(1, 1) <> "SNO" Then
            MsgBox "サンプルナンバーのQCODEに［SNO］以外が設定されています。" & vbCrLf & "修正指示ファイルの［修正ファイル名］を確認してください。", vbExclamation, "MCS 2020 - Indata_Revision"
            End
        End If
        
        ' 修正対象サンプルナンバーを検索、あったら行を取得
        Set FoundCell = Range(Cells(7, 1), ws_data.Cells(max_row, 1)).Find(What:=rev_sno, lookat:=xlWhole)
        If FoundCell Is Nothing Then
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = "指定したサンプルナンバーがデータ上で見つかりません。"
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "修正できませんでした。"
        Else
            dat_row = FoundCell.Row
        End If

        ' 修正対象項目（QCODE）を検索、あったら列を取得
        Set FoundCell = Range(Cells(1, 1), ws_data.Cells(1, max_col)).Find(What:=rev_qcode, lookat:=xlWhole)
        If FoundCell Is Nothing Then
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = "対象のQCODEが見つかりません。"
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "修正できませんでした。"
        Else
            dat_col = FoundCell.Column
            If rev_mact <> 0 Then
                dat_col = dat_col + rev_mact - 1
            End If
        End If

        ' 修正内容（rev_after）が、『ブランク（NULL）』、『クリア』、『無回答』なら
        If (rev_after = "") Or (rev_after = "クリア") Or (rev_after = "無回答") Then
            ' 修正前のデータを出力
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = ws_data.Cells(dat_row, dat_col).Value
            ' データを修正
            ws_data.Cells(dat_row, dat_col).Value = ""
            ' 修正後のデータを出力
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "クリア（無回答）"
            
        ' 修正内容（rev_after）が数値なら
        ElseIf IsNumeric(rev_after) = True Then
            ' 修正前のデータを出力
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = ws_data.Cells(dat_row, dat_col).Value
            ' データを修正
            ws_data.Cells(dat_row, dat_col).Value = rev_after
            ' 修正後のデータを出力
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = ws_data.Cells(dat_row, dat_col).Value
            
        ' 修正内容（rev_after）が"DEL"なら
        ElseIf rev_after = "DEL" Then
            ' データを修正
            ws_data.Cells(dat_row, dat_col).EntireRow.Delete
            ' 修正後のデータを出力
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "サンプルカット"
            
        ' 修正内容（rev_after）が上記以外なら
        Else
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "未処理"
        End If
    Next rev_cnt

    Application.DisplayAlerts = False
    wb_data.SaveAs Filename:="H:\" & gcode & "\MCS\1_DATA\" & odata_fn
    wb_data.Close
    ws_revision.Activate
    ws_revision.Cells(6, 1).Select
    wb_revision.SaveAs Filename:="H:\" & gcode & "\MCS\4_LOG\" & gcode & "RE_log.xlsx"
    wb_revision.Save
    wb_revision.Close
    Application.DisplayAlerts = True

    Set wb_revision = Nothing
    Set ws_revision = Nothing
    
    Application.StatusBar = False
    ws_mainmenu.Activate
    Cells(1, 1).Select
    
' システムログの出力
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "03"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 03"
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
         "\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Output As #1
        Print #1, ws_mainmenu.Cells(gcode_row, gcode_col) & " MCS 2020 operation history"
        Close #1
    End If
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 入力データの修正：使用ファイル［" & rev_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "データの修正が完了しました。", vbInformation, "MCS 2020 - Indata_Revision"
End Sub

