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
'�@SPSS�pCSV�t�@�C���E�V���^�b�N�X�̍쐬�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@ '
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�c���@�`�W�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.05.17�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
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
    ot_fn = Application.GetOpenFilename("�f�[�^�t�@�C��,*.xlsx", , "�f�[�^�t�@�C�����J��")
    
    If InStr(ot_fn, "IN.xlsx") = 0 Then
    
    ElseIf InStr(ot_fn, "OT.xlsx") = 0 Then
    
    ElseIf InStr(ot_fn, "RE.xlsx") = 0 Then
    
    End If
    
    If ot_fn = "False" Then
        ' �L�����Z���{�^���̏���
        End
    ElseIf ot_fn = "" Then
        MsgBox "SPSS�p CSV�t�@�C�����쐬����m�f�[�^�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - SPSScsv_Creation"
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

' ��������SPSS�pCSV�t�@�C���̍쐬�R�[�f�B���O
    ot_fn = Dir(ot_fn)
    Set ot_wb = Workbooks(ot_fn)
    Set ot_ws = ot_wb.Worksheets(1)
    
    ' �����Ώۃf�[�^�t�@�C���̍s�񐔂̎擾
    ot_ws.Activate
    ot_col = ot_ws.Cells(1, Columns.Count).End(xlToLeft).Column
    ot_row = ot_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Application.ScreenUpdating = False
    Open file_path & "\1_DATA\" & ws_mainmenu.Cells(3, 8) & ".sps" For Output As #1
    For i_cnt = 1 To ot_col
        DoEvents
        v_index = Qcode_Match(ot_ws.Cells(1, i_cnt))
        ' �T���v���i���o�[�̏����i�����́m6�n�ŃV���^�b�N�X�o�́j
        If q_data(v_index).q_code = "SNO" Then
            '�_�~�[�w�b�_�[�̏���
            ot_ws.Cells(6, i_cnt) = String(6, "9")
            
            ' �V���^�b�N�X�̏o��
            Print #1, "PRINT    FORMAT SNO (F6)."
            Print #1, "VARIABLE LABELS SNO '�T���v���i���o�['."
            Print #1, "VARIABLE LEVEL SNO (scale)."
        ' *���H�ド�x���̏���
        ElseIf q_data(v_index).q_code = "*���H��" Then
            ot_ws.Cells(1, i_cnt) = "���H��"
        ElseIf q_data(v_index).q_format = "S" Then
            '�_�~�[�w�b�_�[�̏���
            ot_ws.Cells(6, i_cnt) = String(Len(Format(q_data(v_index).ct_count)), "9")
            
            ' �V���^�b�N�X�̏o��
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
        ' �}���`�A���T�[�̏���
        ElseIf (q_data(v_index).q_format = "M") Or (Mid(q_data(v_index).q_format, 1, 1) = "L") Then
            If q_data(v_index).ct_count <> 0 Then
                '�f�[�^�̃w�b�_�[�̏���
                ot_ws.Cells(1, i_cnt) = ot_ws.Cells(1, i_cnt) & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0"))
                
                '�_�~�[�w�b�_�[�̏���
                ot_ws.Cells(6, i_cnt) = "9"
                
                ' �V���^�b�N�X�̏o��
                Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " (F1)."
                Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " '" & q_data(v_index).q_title _
                 & "�F" & q_data(v_index).q_ct(ot_ws.Cells(2, i_cnt)) & "'."
                Print #1, "   VALUE LABELS " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " 1 '�Y��'."
                Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & "_" & Format(ot_ws.Cells(2, i_cnt), String(Len(Format(q_data(v_index).ct_count)), "0")) & " (nominal)."
                
                If Val(ot_ws.Cells(2, i_cnt)) = 1 Then
                    '�u�P�E�O�v�̏���
                    For n_cnt = 7 To ot_row
                        If WorksheetFunction.Sum(Range(ot_ws.Cells(n_cnt, i_cnt), ot_ws.Cells(n_cnt, i_cnt + q_data(v_index).ct_count - 1))) > 0 Then
                            With Range(ot_ws.Cells(n_cnt, i_cnt), ot_ws.Cells(n_cnt, i_cnt + q_data(v_index).ct_count - 1))
                             .Replace What:="", Replacement:="0", lookat:=xlWhole
                            End With
                        End If
                    Next n_cnt
                End If
            End If
        ' ���A���A���T�[�̏����i�g�J�[�\���܂ށj
        ElseIf (Mid(q_data(v_index).q_format, 1, 1) = "R") Or (q_data(v_index).q_format = "H") Then
            '�_�~�[�w�b�_�[�̏���
            ot_ws.Cells(6, i_cnt) = String(q_data(v_index).r_byte, "9")
            
            ' �V���^�b�N�X�̏o��
            Print #1, "PRINT    FORMAT " & q_data(v_index).q_code & " (F" & q_data(v_index).r_byte & ")."
            Print #1, "VARIABLE LABELS " & q_data(v_index).q_code & " '" & q_data(v_index).q_title & "'."
            Print #1, "VARIABLE LEVEL " & q_data(v_index).q_code & " (scale)."
        ' �t���[�A���T�[�̏���
        ElseIf (q_data(v_index).q_format = "F") Or (q_data(v_index).q_format = "O") Then
            '�_�~�[�w�b�_�[�̏���
            ot_ws.Cells(6, i_cnt) = String(255, "*")
            
            ' �V���^�b�N�X�̏o��
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
    
' �V�X�e�����O�̏o��
    ' 2020.6.3 - �ǉ�
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - SPSS�pCSV�t�@�C���A�V���^�b�N�X�t�@�C���̍쐬�F�Ώۃt�@�C���m" & ot_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "SPSS�pCSV�t�@�C���ƃV���^�b�N�X�t�@�C�����o�͂��܂����B", vbInformation, "MCS 2020 - SPSScsv_Creation"
End Sub

