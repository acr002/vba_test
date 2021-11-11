Attribute VB_Name = "Module05"
Option Explicit
    Dim wb_error As Workbook
    Dim ws_error As Worksheet
    Dim msg_flg As Integer
    Dim ok_flg As Integer, pass_flg As Integer

Sub Data_Check(ope_code As String)
    Dim temp_data As Variant
    Dim waitTime As Variant
    
    Dim data_fn As String
    Dim period_pos As Integer
    Dim set_row As Long, set_col As Long
    Dim max_row As Long, max_col As Long
    Dim r_cnt As Long, c_cnt As Long, i_cnt As Long, m_cnt As Long, q_cnt As Long
    Dim d_index As Long
    Dim err_row As Long
    Dim ma_qcode As String
    Dim ra_data  As Variant
    Dim ra_len As Long, ra_int As Long
    Dim ra_mod As Double

    ' �G���[���X�g�̏o�̓A�h���X�ݒ�
    Const err_sno As Integer = 1      ' �T���v���ԍ�
    Const err_qcode As Integer = 2    ' QCODE
    Const err_msg As Integer = 4      ' �G���[���e
    Const err_data As Integer = 5     ' �G���[�f�[�^
    Const err_rst  As Integer = 6     ' �C����
'--------------------------------------------------------------------------------------------------'
'�@���̓f�[�^�̃��W�b�N�`�F�b�N�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.27�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Setup_Hold    ' �ǉ�
    Call Filepath_Get
    Call Setup_Check
    
    Open file_path & "\4_LOG\" & ope_code & "err.xlsx" For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(ope_code & "err.xlsx").Close
    End If
    
    If Dir(file_path & "\4_LOG\" & ope_code & "err.xlsx") <> "" Then
        Kill file_path & "\4_LOG\" & ope_code & "err.xlsx"
    End If
    
    ChDrive file_path & "\1_DATA"
    ChDir file_path & "\1_DATA"
    
    wb.Activate
    ws_mainmenu.Select
    data_fn = Application.GetOpenFilename("�f�[�^�t�@�C��,*.xlsx", , "�f�[�^�t�@�C�����J��")
    If data_fn = "False" Then
        ' �L�����Z���{�^���̏���
        ws_mainmenu.Select
        End
    ElseIf data_fn = "" Then
        MsgBox "�`�F�b�N����m�f�[�^�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Data_Check"
        End
    End If
    
    ' �t���p�X����t�@�C�����̎擾
    data_fn = Dir(data_fn)
    
    ' �m�[�G���[�t�@�C�����������ꍇ�͍폜
    If Dir(file_path & "\4_LOG\" & data_fn & "_No Error") <> "" Then
        Kill file_path & "\4_LOG\" & data_fn & "_No Error"
    End If
    
    wb.Activate
    ws_mainmenu.Select

    If Dir(file_path & "\1_DATA\" & data_fn) <> "" Then
        outdata_fn = data_fn
        status_msg = "�`�F�b�N�Ώۃt�@�C���m" & data_fn & "�n �I�[�v����..."
        Call Datafile_Open
        Call Setup_Hold
        wb_outdata.Activate
        ws_outdata.Select
        
        If ActiveSheet.AutoFilterMode = True Then
            wb.Activate
            ws_mainmenu.Select
            wb_outdata.Activate
            ws_outdata.Select
            MsgBox "�f�[�^�t�@�C���̃I�[�g�t�B���^���������Ă��������B", vbExclamation, "MCS 2020 - Data_Check"
            Application.StatusBar = False
            End
        End If
        
        max_col = ws_outdata.Cells(1, Columns.Count).End(xlToLeft).Column
        max_row = ws_outdata.Cells(Rows.Count, setup_col).End(xlUp).Row
        For i_cnt = max_row To Rows.Count
            If WorksheetFunction.CountA(Rows(i_cnt)) = 0 Then
                max_row = i_cnt - 1
                Exit For
            End If
        Next i_cnt
        Application.StatusBar = False
    Else
        MsgBox "�t�@�C�� " & data_fn & " �����݂��܂���B", vbExclamation, "MCS 2020 - Data_Check"
        Application.StatusBar = False
        End
    End If

    msg_flg = 1
    pass_flg = 0
    err_row = 1

    wb.Activate
    ws_setup.Select
    set_col = Cells(1, Columns.Count).End(xlToLeft).Column
    set_row = Cells(Rows.Count, setup_col).End(xlUp).Row

    ' �ݒ��ʏ��ƃf�[�^���C�A�E�g�̃`�F�b�N
    For i_cnt = 3 To set_row
        DoEvents
        Application.StatusBar = "�ݒ��ʂƃf�[�^�t�@�C���̃��W�b�N�`�F�b�N�� ... " & Int(i_cnt / set_row * 100) & "%/100%"
        If pass_flg <> 1 Then
            ok_flg = 0
            ' QCODE�̑��݃`�F�b�N
            For c_cnt = 1 To max_col
                If ws_setup.Cells(i_cnt, 1) = ws_outdata.Cells(1, c_cnt) Then
                    ok_flg = 1
                    Exit For
                End If
            Next c_cnt
            If ok_flg = 0 Then
                If ws_setup.Cells(i_cnt, 1) = "*���H��" Then
                    pass_flg = 1
                ElseIf Mid(ws_setup.Cells(i_cnt, 1), 1, 1) <> "*" Then
                    wb_outdata.Activate
                    ws_outdata.Select
                    MsgBox "�ݒ��ʂɂ���QCODE�m" & ws_setup.Cells(i_cnt, 1) & "�n���A�f�[�^��ɂ���܂���B" & vbCrLf & "�`�F�b�N����f�[�^�̃t�@�C�����m�F���Ă��������B", vbCritical, "MCS 2020 - Data_Check"
                    Application.StatusBar = False
                    ws_outdata.Cells(1, 1).Select
                    End
                End If
            End If

            ' �ݖ�`���ɂ��`�F�b�N
            If Mid(ws_setup.Cells(i_cnt, 1), 1, 1) <> "*" Then
                If ws_setup.Cells(i_cnt, 9) = "S" Then
                    q_cnt = 0
                    For c_cnt = 1 To max_col
                        If ws_setup.Cells(i_cnt, 1) = ws_outdata.Cells(1, c_cnt) Then
                            q_cnt = q_cnt + 1
                        End If
                    Next c_cnt
                    If q_cnt <> 1 Then
                        wb_outdata.Activate
                        ws_outdata.Select
                        MsgBox "�V���O���A���T�[��QCODE�m" & ws_setup.Cells(i_cnt, 1) & "�n���A�f�[�^��� " & Trim(Str(q_cnt)) & "�� ����܂��B" & vbCrLf & "�`�F�b�N����f�[�^�̃t�@�C�����m�F���Ă��������B", vbCritical, "MCS 2020 - Data_Check"
                        Application.StatusBar = False
                        End
                    End If
                ElseIf ws_setup.Cells(i_cnt, 9) = "R" Then
                    q_cnt = 0
                    For c_cnt = 1 To max_col
                        If ws_setup.Cells(i_cnt, 1) = ws_outdata.Cells(1, c_cnt) Then
                            q_cnt = q_cnt + 1
                        End If
                    Next c_cnt
                    If q_cnt <> 1 Then
                        wb_outdata.Activate
                        ws_outdata.Select
                        MsgBox "���A���A���T�[��QCODE�m" & ws_setup.Cells(i_cnt, 1) & "�n���A�f�[�^��� " & Trim(Str(q_cnt)) & "�� ����܂��B" & vbCrLf & "�`�F�b�N����f�[�^�̃t�@�C�����m�F���Ă��������B", vbCritical, "MCS 2020 - Data_Check"
                        Application.StatusBar = False
                        End
                    End If
                ElseIf ws_setup.Cells(i_cnt, 9) = "H" Then
                    q_cnt = 0
                    For c_cnt = 1 To max_col
                        If ws_setup.Cells(i_cnt, 1) = ws_outdata.Cells(1, c_cnt) Then
                            q_cnt = q_cnt + 1
                        End If
                    Next c_cnt
                    If q_cnt <> 1 Then
                        wb_outdata.Activate
                        ws_outdata.Select
                        MsgBox "�g�J�[�\����QCODE�m" & ws_setup.Cells(i_cnt, 1) & "�n���A�f�[�^��� " & Trim(Str(q_cnt)) & "�� ����܂��B" & vbCrLf & "�`�F�b�N����f�[�^�̃t�@�C�����m�F���Ă��������B", vbCritical, "MCS 2020 - Data_Check"
                        Application.StatusBar = False
                        End
                    End If
                ElseIf (ws_setup.Cells(i_cnt, 9) = "M") Or (Mid(ws_setup.Cells(i_cnt, 9), 1, 1) = "L") Then
                    q_cnt = 0
                    For c_cnt = 1 To max_col
                        If ws_setup.Cells(i_cnt, 1) = ws_outdata.Cells(1, c_cnt) Then
                            q_cnt = q_cnt + 1
                        End If
                    Next c_cnt
                    If q_cnt <> Val(ws_setup.Cells(i_cnt, 16)) Then
                        wb_outdata.Activate
                        ws_outdata.Select
                        MsgBox "�}���`�A���T�[��QCODE�m" & ws_setup.Cells(i_cnt, 1) & "�n��CT���ƃ��C�A�E�g�̗񐔂���v���܂���B" & vbCrLf & vbCrLf & "CT���m" & ws_setup.Cells(i_cnt, 16) & "�n�@�񐔁m" & q_cnt & "�n" & vbCrLf & vbCrLf & "�`�F�b�N����f�[�^�̃t�@�C�����m�F���Ă��������B", vbCritical, "MCS 2020 - Data_Check"
                        Application.StatusBar = False
                        End
                    End If
                End If
          End If
        End If
    Next i_cnt
    Application.StatusBar = False
    
    ' �T���v�����̊e�ݖ�̃`�F�b�N
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - �f�[�^�̃��W�b�N�`�F�b�N"
    Form_Progress.Repaint
    progress_msg = "�f�[�^�̃��W�b�N�`�F�b�N���L�����Z�����܂����B"
    Application.Visible = False
    AppActivate Form_Progress.Caption
    
    ma_qcode = ""
    For c_cnt = 1 To max_col
        DoEvents
        Form_Progress.Label1.Caption = Int(c_cnt / max_col * 100) & "%"
        Form_Progress.Label2.Caption = data_fn & " �̐ݖ⍀�ڃ��W�b�N�`�F�b�N��" & Status_Dot(c_cnt)

        d_index = Qcode_Match(ws_outdata.Cells(1, c_cnt))
'        temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column))

        Select Case Left(q_data(d_index).q_format, 1)
        Case "S"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column))
            For r_cnt = 1 To max_row - 6
                If Val(temp_data(r_cnt, 1)) > Val(q_data(d_index).ct_count) Then
                    Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                    ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                    ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                    ws_error.Cells(err_row, err_msg) = "�����W�I�[�o�["
                    ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                    ws_error.Cells(err_row, err_rst) = "�N���A"
                    err_row = err_row + 1
                End If
            Next r_cnt
        Case "M"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column + q_data(d_index).ct_loop - 1))
            If ma_qcode <> ws_outdata.Cells(1, q_data(d_index).data_column) Then
                For m_cnt = 1 To q_data(d_index).ct_loop
                    For r_cnt = 1 To max_row - 6
                        If (Val(temp_data(r_cnt, m_cnt)) <> 1) And (Val(temp_data(r_cnt, m_cnt)) <> 0) Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_qcode + 1) = m_cnt
                            ws_error.Cells(err_row, err_msg) = "�}���`�A���T�[�Łm1�n�ȊO�����͂���Ă��܂��B"
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, m_cnt)
                            ws_error.Cells(err_row, err_rst) = 1
                            err_row = err_row + 1
                        End If
                    Next r_cnt
                Next m_cnt
                ma_qcode = ws_outdata.Cells(1, q_data(d_index).data_column)
            End If
        Case "L"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column + q_data(d_index).ct_loop - 1))
            If ma_qcode <> ws_outdata.Cells(1, q_data(d_index).data_column) Then
                For m_cnt = 1 To q_data(d_index).ct_loop
                    For r_cnt = 1 To max_row - 6
                        If (Val(temp_data(r_cnt, m_cnt)) <> 1) And (Val(temp_data(r_cnt, m_cnt)) <> 0) Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_qcode + 1) = m_cnt
                            ws_error.Cells(err_row, err_msg) = "���~�b�g�}���`�A���T�[�Łm1�n�ȊO�����͂���Ă��܂��B"
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, m_cnt)
                            ws_error.Cells(err_row, err_rst) = 1
                            err_row = err_row + 1
                        End If
                    Next r_cnt
                Next m_cnt
                ma_qcode = ws_outdata.Cells(1, q_data(d_index).data_column)
            End If
        Case "R"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column))
            For r_cnt = 1 To max_row - 6
                If temp_data(r_cnt, 1) <> "" Then
                    '�f�[�^�̃T�C�Y�`�F�b�N
                    If Val(temp_data(r_cnt, 1)) > 2147483647 Then
                        Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                        ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                        ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                        ws_error.Cells(err_row, err_msg) = "�I�[�o�[�t���E"
                        ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                        ws_error.Cells(err_row, err_rst) = "�N���A"
                        err_row = err_row + 1
                    Else
                        ra_data = Val(temp_data(r_cnt, 1))
                        ra_len = Len(temp_data(r_cnt, 1))
                        ra_mod = Val(temp_data(r_cnt, 1))
                        ra_int = Int(ra_mod)

                        '���`�F�b�N
                        If Len(ra_data) > q_data(d_index).r_byte Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_msg) = "���I�[�o�["
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                            ws_error.Cells(err_row, err_rst) = "�N���A"
                            err_row = err_row + 1
                        End If

                        '�͈͋L���`�F�b�N
                        If ra_len <> Len(ra_data) Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_msg) = "�f�[�^���m�F���Ă��������B"
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                            err_row = err_row + 1
                        End If

                        '�}�C�i�X�l�`�F�b�N
                        If ra_data < 0 Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_msg) = "�}�C�i�X�f�[�^"
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                            err_row = err_row + 1
                        End If

                        '�����_�`�F�b�N
                        If ra_int <> ra_mod Then
                            Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                            ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                            ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                            ws_error.Cells(err_row, err_msg) = "�����_����"
                            ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                            ws_error.Cells(err_row, err_rst) = Format(temp_data(r_cnt, 1), "0")
                            err_row = err_row + 1
                        End If
                    End If
                End If
            Next r_cnt
        Case "H"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column))
            For r_cnt = 1 To max_row - 6
                If Val(temp_data(r_cnt, 1)) > 100 Then
                    Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                    ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                    ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                    ws_error.Cells(err_row, err_msg) = "�g�J�[�\���m100�n�I�[�o�["
                    ws_error.Cells(err_row, err_data) = temp_data(r_cnt, 1)
                    ws_error.Cells(err_row, err_rst) = 100
                    err_row = err_row + 1
                End If
            Next r_cnt
        Case "C"
            temp_data = Range(ws_outdata.Cells(7, q_data(d_index).data_column), ws_outdata.Cells(max_row, q_data(d_index).data_column))
            ' �T���v���i���o�[�̃`�F�b�N
            If ws_outdata.Cells(1, c_cnt) = "SNO" Then
                For r_cnt = 1 To max_row - 6
                    DoEvents
                    Form_Progress.Label1.Caption = Int(r_cnt / (max_row - 6) * 100) & "%"
                    Form_Progress.Label2.Caption = data_fn & " �̃T���v���i���o�[�`�F�b�N��..."

' �������d���̂ŁA�R�����g�A�E�g���܂����B
' �T���v���i���o�[�̏d���`�F�b�N�́A�ǂ����ŏ������������悢�Ǝv���܂��c�i�L�E�ցE�M�j
'                    If WorksheetFunction.CountIf(Range(ws_outdata.Cells(7, c_cnt), ws_outdata.Cells(max_row, c_cnt)), ws_outdata.Cells(r_cnt + 6, c_cnt)) > 1 Then
'                        Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
'                        ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
'                        ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
'                        ws_error.Cells(err_row, err_msg) = "�T���v���ԍ����d�����Ă��܂��B"
'                        ws_error.Cells(err_row, err_data) = Format(r_cnt + 6) & "�s��"
'                        ws_error.Cells(err_row, err_rst) = "�d�����Ȃ��悤�ɏC�����Ă��������"
'                        err_row = err_row + 1
'                    End If
                    
                    If temp_data(r_cnt, 1) = "" Then
                        Call Error_Message(err_row, err_sno, err_qcode, err_msg, err_data, err_rst)
                        ws_error.Cells(err_row, err_sno) = ws_outdata.Cells(r_cnt + 6, 1)
                        ws_error.Cells(err_row, err_qcode) = ws_outdata.Cells(1, c_cnt)
                        ws_error.Cells(err_row, err_msg) = "�T���v���ԍ������񓚂ł��B"
                        ws_error.Cells(err_row, err_data) = Format(r_cnt + 6) & "�s��"
                        ws_error.Cells(err_row, err_rst) = "���񓚂��Ȃ��悤�ɏC�����Ă��������"
                        err_row = err_row + 1
                    End If
                Next r_cnt
            End If
        End Select
    Next c_cnt
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")
    Application.Wait waitTime
    Application.Visible = True
    Unload Form_Progress

    If Not wb_error Is Nothing Then
        wb_error.Activate
        ws_error.Select
        ws_error.Cells.Select
        With Selection.Font
            .Name = "Takao�S�V�b�N"
            .Size = 11
        End With
        
        Columns("D:D").Select
        Columns("D:D").EntireColumn.AutoFit
        ws_error.Cells(1, 1).Select
        
        ' �G���[���b�Z�[�W�̃\�[�e�B���O����
        ActiveSheet.Sort.SortFields.Clear
        
        ActiveSheet.Sort.SortFields.Add _
        Key:=Range("A1"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        
        ActiveSheet.Sort.SortFields.Add _
        Key:=Range("A2"), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
        
        With ActiveSheet.Sort
            .SetRange Columns("A:F")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        If Dir(file_path & "\4_LOG\" & ope_code & "err.xlsx") = "" Then
            wb_error.SaveAs Filename:=file_path & "\4_LOG\" & ope_code & "err.xlsx"
        Else
            Kill file_path & "\4_LOG\" & ope_code & "err.xlsx"
            wb_error.SaveAs Filename:=file_path & "\4_LOG\" & ope_code & "err.xlsx"
        End If
    End If

    ' �f�[�^�t�@�C���̃N���[�Y
    wb_outdata.Activate
    ws_outdata.Select
    Range("B7").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
    Set wb_error = Nothing
    Set ws_error = Nothing
    
    wb.Activate
    ws_mainmenu.Select
    
' �V�X�e�����O�̏o��
    ' 2020.6.3 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = ope_code
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > " & ope_code
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �f�[�^�̃��W�b�N�`�F�b�N�F�Ώۃt�@�C���m" & data_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�f�[�^�t�@�C���̃��W�b�N�`�F�b�N���������܂����B", vbInformation, "MCS 2020 - Data_Check"
    
    ' ���W�b�N�`�F�b�N��ɁA�G���[���O�t�@�C�����Ȃ���΁A�m�[�G���[�t�@�C�����쐬����
    If Dir(file_path & "\4_LOG\" & ope_code & "err.xlsx") = "" Then
        Open file_path & "\4_LOG\" & data_fn & "_No Error" For Append As #1
        Close #1
    End If
End Sub

Private Sub Error_Message(ByRef e_row As Long, ByVal e_sno As Integer, _
 ByVal e_qcode As Integer, ByVal e_msg As Integer, _
 ByVal e_data As Integer, ByVal e_rst As Integer)
' �G���[���b�Z�[�W�o�̓t�@�C���̍쐬
    If msg_flg = 1 Then
        msg_flg = 0
        Workbooks.Add
        Set wb_error = ActiveWorkbook
        Set ws_error = wb_error.Worksheets("Sheet1")
        
        ws_error.Range("A1:F1").Select
        With Selection
            .HorizontalAlignment = xlHAlignCenter
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(58, 56, 56)
        End With
        Rows(1).RowHeight = 18
        Range("A:C").EntireColumn.ColumnWidth = 8.5
        Columns("D:D").ColumnWidth = 49.88
        Range("E:F").EntireColumn.ColumnWidth = 8.88
        
        ws_error.Cells(e_row, e_sno) = "SampleNo"
        ws_error.Cells(e_row, e_qcode) = "QCODE"
        ws_error.Cells(e_row, e_qcode + 1) = "MA_CT"
        ws_error.Cells(e_row, e_msg) = "�G���[���e"
        ws_error.Cells(e_row, e_data) = "�񓚓��e"
        ws_error.Cells(e_row, e_rst) = "�C����"
        e_row = e_row + 1
    End If
End Sub

