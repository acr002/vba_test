Attribute VB_Name = "Module32"
Option Explicit
    Dim ot_wb As Workbook
    Dim ot_ws As Worksheet

    Dim rulz_one As Object, rulz_two As Object, rulz_thr As Object
    Dim msg_row As Variant, msg_col As Variant, lbl_nm As String

    Dim rd_fn As String, ot_fn As String, newworkbook_fn As String, qcode As String
    Dim alt_msg As String, alrt_msg As String, alt_nm As String

    Dim s_r As Long, ma_ed As Long
    Dim j As Long, code_cnt As Long, ct_cnt As Long

    Dim adr_arr As Variant

    Dim is_num As Boolean

Public Sub SPSScsv_Creation()
    Dim i_cnt As Long, c_cnt As Long, n_cnt As Long
    Dim taget_fn As String
    Dim ot_row As Long
    Dim ot_col As Long
    Dim v_index As Long
    Dim val_label As String
'--------------------------------------------------------------------------------------------------'
'　SPSS用CSVファイル・シンタックスの作成　　　　　　　　　　　　　　　　　　　　　　　　　　　　　 '
'--------------------------------------------------------------------------------------------------'
'　作成者　　田中　義晃　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.05.17　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Hold
    Call Setup_Check
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\1_DATA"
    
    wb.Activate
    ws_mainmenu.Select
    rd_fn = ws_mainmenu.Cells(gcode_row, gcode_col) & "SPSS.csv"
    
step00:
    ot_fn = Application.GetOpenFilename("データファイル,*.xlsx", , "データファイルを開く")
    
    If InStr(ot_fn, "IN.xlsx") = 0 Then
    
    ElseIf InStr(ot_fn, "OT.xlsx") = 0 Then
    
    ElseIf InStr(ot_fn, "RE.xlsx") = 0 Then
    
    End If
    
    If ot_fn = "False" Then
        ' キャンセルボタンの処理
        End
    ElseIf ot_fn = "" Then
        MsgBox "SPSS用 CSVファイルを作成する［データファイル］を選択してください。", vbExclamation, "MCS 2020 - SPSScsv_Creation"
        Application.StatusBar = False
        GoTo step00
    End If
    
    Open ot_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(Dir(ot_fn)).Close
    Else
        Workbooks.Open ot_fn
    End If

    Open file_path & "\1_DATA\" & rd_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(rd_fn).Close
    End If

' ここからSPSS用CSVファイルの作成コーディング
    ot_fn = Dir(ot_fn)
    Set ot_wb = Workbooks(ot_fn)
    Set ot_ws = ot_wb.Worksheets(1)
    
    ' 処理対象データファイルの行列数の取得
    ot_ws.Activate
    ot_col = ot_ws.Cells(1, Columns.Count).End(xlToLeft).Column
    ot_row = ot_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    Open file_path & "\1_DATA\" & ws_mainmenu.Cells(3, 8) & ".sps" For Output As #1
    For i_cnt = 1 To ot_col
        DoEvents
        v_index = Qcode_Match(ot_ws.Cells(1, i_cnt))
        ' サンプルナンバーの処理（長さは［6］でシンタックス出力）
        If q_data(v_index).q_code = "SNO" Then
            'ダミーヘッダーの処理
            ot_ws.Cells(6, i_cnt) = String(6, "9")
            
            ' シンタックスの出力
            Print #1, "PRINT    FORMAT SNO (F6)."
            Print #1, "VARIABLE LABELS SNO 'サンプルナンバー'."
            Print #1, "VARIABLE LEVEL SNO (scale)."
        ' *加工後ラベルの処理
        ElseIf q_data(v_index).q_code = "*加工後" Then
            ot_ws.Cells(1, i_cnt) = "加工後"
        ElseIf q_data(v_index).q_format = "S" Then
            'ダミーヘッダーの処理
            ot_ws.Cells(6, i_cnt) = String(Len(Format(q_data(v_index).ct_count)), "9")
            
            ' シンタックスの出力
            If q_data(v_index).ct_count = 0 Then
                Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & " (F)."
            Else
                Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & " (F" & Len(Format(q_data(v_index).ct_count)) & ")."
            End If
            Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & " '" & q_data(v_index).q_title & "'."
            If q_data(v_index).ct_count <> 0 Then
                val_label = ""
                For c_cnt = 1 To q_data(v_index).ct_count
                    val_label = val_label & " " & c_cnt & " '" & q_data(v_index).q_ct(c_cnt) & "'"
                Next c_cnt
                Print #1, "   VALUE LABELS " & q_data(v_index).q_code & val_label & "."
                Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & " (nominal)."
            End If
        ' マルチアンサーの処理
        ElseIf (q_data(v_index).q_format = "M") Or (Mid(q_data(v_index).q_format, 1, 1) = "L") Then
            If q_data(v_index).ct_count <> 0 Then
                'データのヘッダーの処理
                ot_ws.Cells(1, i_cnt) = ot_ws.Cells(1, i_cnt) & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0"))
                
                'ダミーヘッダーの処理
                ot_ws.Cells(6, i_cnt) = "9"
                
                ' シンタックスの出力
                Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " (F1)."
                Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " '" & q_data(v_index).q_title _
                 & "：" & q_data(v_index).q_ct(ot_ws.Cells(2, i_cnt)) & "'."
                Print #1, "   VALUE LABELS " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " 1 '該当'."
                Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " (nominal)."
                
                If Val(ot_ws.Cells(2, i_cnt)) = 1 Then
                    '「１・０」の処理
                    For n_cnt = 7 To ot_row
                        If WorksheetFunction.Sum(Range(ot_ws.Cells(n_cnt, i_cnt), ot_ws.Cells(n_cnt, i_cnt + q_data(v_index).ct_count - 1))) > 0 Then
                            With Range(ot_ws.Cells(n_cnt, i_cnt), ot_ws.Cells(n_cnt, i_cnt + q_data(v_index).ct_count - 1))
                             .Replace What:="", Replacement:="0", lookat:=xlWhole
                            End With
                        End If
                    Next n_cnt
                End If
            End If
        ' リアルアンサーの処理（Ｈカーソル含む）
        ElseIf (Mid(q_data(v_index).q_format, 1, 1) = "R") Or (q_data(v_index).q_format = "H") Then
            'ダミーヘッダーの処理
            ot_ws.Cells(6, i_cnt) = String(q_data(v_index).r_byte, "9")
            
            ' シンタックスの出力
            Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & " (F" & q_data(v_index).r_byte & ")."
            Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & " '" & q_data(v_index).q_title & "'."
            Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & " (scale)."
        ' フリーアンサーの処理
        ElseIf (q_data(v_index).q_format = "F") Or (q_data(v_index).q_format = "O") Then
            'ダミーヘッダーの処理
            ot_ws.Cells(6, i_cnt) = String(255, "*")
            
            ' シンタックスの出力
            Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & " (A255)."
            Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & " '" & q_data(v_index).q_title & "'."
            Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & " (nominal)."
        End If
    Next i_cnt
    Close #1
    
    Application.DisplayAlerts = False
    ot_wb.Activate
    Rows("2:5").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ot_wb.SaveAs Filename:=file_path & "\1_DATA\" & rd_fn, FileFormat:=xlCSV, CreateBackup:=False
    ot_wb.Close
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
    
    Set ot_wb = Nothing
    Set ot_ws = Nothing
    
' システムログの出力
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "12"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 12"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - SPSS用CSVファイル、シンタックスファイルの作成：対象ファイル［" & ot_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "SPSS用CSVファイルとシンタックスファイルを出力しました。", vbInformation, "MCS 2020 - SPSScsv_Creation"
End Sub

