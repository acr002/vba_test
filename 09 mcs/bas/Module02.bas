Attribute VB_Name = "Module02"
Option Explicit
    Dim inlayout_wb As Workbook
    Dim inlayout_fn, newworkbook_fn, form_type As String
    Dim inlayout_ws As Worksheet
    Dim s_row_count, s_ct_count, i_col_count
    Dim i, j As Integer

Sub Inlayout_Creation()
'--------------------------------------------------------------------------------------------------'
'　入力レイアウトの作成　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　田中義晃　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.11　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check

    wb.Activate
    ws_mainmenu.Select

    inlayout_fn = Cells(gcode_row, gcode_col) & " 入力レイアウト.xlsx"

    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & inlayout_fn
    If Err.Number <> 0 Then
        ActiveWorkbook.Close
        Open file_path & "\3_FD\" & inlayout_fn For Append As #1
        Close #1
        If Err.Number = 70 Then
            MsgBox inlayout_fn & "はすでに開かれています。" _
             & vbCrLf & "ファイルを閉じてから再実行して下さい。", vbExclamation, "MCS 2020 - Inlayout_Creation"
        End If
        End
    End If

' ここから入力レイアウトの作成コーディング(´･ω･`)
    Set inlayout_wb = ActiveWorkbook
    Set inlayout_ws = inlayout_wb.Worksheets(1)
    inlayout_ws.Rows(1).NumberFormat = "@"
    s_row_count = ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
    i_col_count = 0
    For i = 3 To s_row_count
        If ws_setup.Cells(i, 1).Value = "*加工後" Then
            Exit For
        End If
        s_ct_count = ws_setup.Cells(i, 16).Value
        If s_ct_count = 0 Then
            s_ct_count = 1
        End If
        form_type = Left(ws_setup.Cells(i, 9).Value, 1)
        Select Case form_type
            Case "M", "L", "F"
                For j = 1 To s_ct_count
                    Call col_set
                    Call bg_set
                Next j
            Case "C", "S", "R", "H"
                Call col_set
                Call bg_set
            Case "O"
                Call col_set
            Case Else
        End Select
    Next i

    inlayout_ws.Range(Cells(1, 1), Cells(6, i_col_count)).Font.Name = ws_setup.Cells(3, 1).Font.Name
    With inlayout_ws.Range(Cells(1, 1), Cells(6, i_col_count))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    With inlayout_ws.Range(Cells(2, 1), Cells(4, i_col_count))
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With inlayout_ws.Range(Cells(5, 1), Cells(6, i_col_count))
        .Borders(xlInsideHorizontal).LineStyle = xlDot
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
    inlayout_ws.Cells(7, 2).Activate
    ActiveWindow.FreezePanes = True
    inlayout_wb.Save
    Set inlayout_wb = Nothing
    Set inlayout_ws = Nothing
    
' システムログの出力
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "01"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 01"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 入力データレイアウト［" & inlayout_fn & "］の作成"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "入力レイアウトが完成しました。", vbInformation, "MCS 2020 - Inlayout_Creation"
End Sub

Private Sub col_set()
    Dim max_range As String
    Dim k As Integer
    Dim fc_o As FormatConditions
    Dim fc_jk As String
    i_col_count = i_col_count + 1
    inlayout_ws.Cells(1, i_col_count).Value = ws_setup.Cells(i, 1).Value
    Select Case form_type
        Case "C"
            If i_col_count = 1 Then
                inlayout_ws.Cells(5, i_col_count).Value = "Low"
                inlayout_ws.Cells(6, i_col_count).Value = "High"
            End If
        Case "S"
            inlayout_ws.Cells(4, i_col_count).Value = "SA"
            inlayout_ws.Cells(5, i_col_count).Value = 1
            inlayout_ws.Cells(6, i_col_count).Value = ws_setup.Cells(i, 16).Value
        Case "M"
            inlayout_ws.Cells(2, i_col_count).Value = j
            inlayout_ws.Cells(4, i_col_count).Value = "MA"
        Case "L"
            inlayout_ws.Cells(2, i_col_count).Value = j
            inlayout_ws.Cells(4, i_col_count).Value = ws_setup.Cells(i, 9).Value
        Case "R"
            For k = 1 To Val(Mid(ws_setup.Cells(i, 9).Value, 2, Len(ws_setup.Cells(i, 9).Value) - 1))
                max_range = max_range & "9"
            Next k
            inlayout_ws.Cells(4, i_col_count).Value = "RA"
            inlayout_ws.Cells(6, i_col_count).Value = max_range
        Case "H"
            inlayout_ws.Cells(4, i_col_count).Value = "HC"
            inlayout_ws.Cells(5, i_col_count).Value = 0
            inlayout_ws.Cells(6, i_col_count).Value = 100
        Case "F"
            With inlayout_ws.Cells(2, i_col_count)
                .Value = ws_setup.Cells(i, 18 + j).Value
                .ShrinkToFit = True
            End With
            inlayout_ws.Cells(4, i_col_count).Value = "FA"
        Case "O"
            inlayout_ws.Cells(4, i_col_count).Value = "FA"
            Set fc_o = ws_setup.Cells(i, 9).FormatConditions
            For k = 1 To fc_o.Count
                fc_jk = Mid(fc_o(k).Formula1, 22, 1)
                If fc_jk = "F" Then
                    inlayout_ws.Range(Cells(1, i_col_count), Cells(6, i_col_count)).Interior.Color = fc_o(k).Interior.Color
                    inlayout_ws.Range(Cells(1, i_col_count), Cells(6, i_col_count)).Font.Color = fc_o(k).Font.Color
                    Exit For
                End If
            Next k
        Case Else
    End Select
End Sub

Private Sub bg_set()
'--------------------------------------------------------------------------------------------------'
'　作成者　田中義晃　　　　　　　　　　　　　　　　　　　　　　　　　作成日　２０１７．０４．１１　'
'--------------------------------------------------------------------------------------------------'
'Dim fc As Object
'Dim ci As Integer
'tana's
    inlayout_ws.Range(Cells(1, i_col_count), Cells(6, i_col_count)).Interior.Color = ws_setup.Cells(i, 9).DisplayFormat.Interior.Color
    inlayout_ws.Range(Cells(1, i_col_count), Cells(6, i_col_count)).Font.Color = ws_setup.Cells(i, 9).DisplayFormat.Font.Color
End Sub

