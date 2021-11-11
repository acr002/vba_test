Attribute VB_Name = "Module54"
Option Explicit
    Dim spread_fn As String
    Dim spread_fd As String
    
    Dim wb_spread As Workbook
    Dim ws_spread0 As Worksheet
    Dim ws_spread1 As Worksheet
    Dim ws_spread2 As Worksheet
    Dim ws_spread3 As Worksheet
    
    Dim wb_print As Workbook
    Dim tn_tab As Boolean

Public Sub Print_spreadsheet()
    Dim waitTime As Variant
    
    Dim rc As Integer
    Dim yen_pos As Long
    
    Dim m_area As Range, tab_no_fc As Range
    Dim head_split As Boolean

    Dim st_rw As Long, ed_rw As Long, usd_rw As Long, gt_rw As Long, _
    dmy_rw As Long, del_rw As Long, top_rw As Long, del_cl As Long, _
    st_cl As Long, ed_cl As Long, height_lim As Long, height_sum As Long, _
    wrap_cnt As Long, ctz_cnt As Long, cpy_cnt As Long, top_ctz As Long, _
    cpy As Long, wrap As Long, del_st As Long, del_ed As Long, s As Long, t As Long

    Dim rng_add As String, face_label As String, print_label As String, _
    table_label As String, not_found As String

    Dim max_cnt As Long, hyo_cnt As Long

    Dim p_cnt As Long, p_temp As Long
    Dim max_row As Long
'--------------------------------------------------------------------------------------------------'
'�@����p�W�v�\�t�@�C���̍쐬  �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�c���`�W�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.04.04�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check

    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\5_INI"
    
    If Dir(file_path & "\5_INI\" & ws_mainmenu.Cells(3, 8) & "cov.xlsx") = "" Then
        MsgBox "�\���e���v���[�g�t�@�C���m*cov.xlsx�n���݂���܂���B" _
         & vbCrLf & "5_INI�t�H���_���ɕ\���e���v���[�g�t�@�C����p�ӂ��Ă��������B", vbExclamation, "MCS 2020 - Cover_Procedure"
        Call Finishing_Mcs2017
        End
    End If
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    rc = MsgBox("����ݒ������W�v�\Excel�t�@�C�����K�v�ł��B" _
     & vbCrLf & "�W�v�\Excel�t�@�C���͂���܂����H" & vbCrLf & vbCrLf _
     & "�u�͂��v�@�� ���łɂ���W�v�\Excel�t�@�C����I��" & vbCrLf & "�u�������v�� �W�v�\Excel�t�@�C�����쐬����", _
     vbYesNoCancel + vbQuestion, "�W�v�\Excel�t�@�C���쐬�̊m�F")
    If rc = vbNo Then
'�E�U�C��ŕ\�����Ȃ��悤�ɂ��� - 2018/6/21
'        MsgBox "�W�v�\Excel�t�@�C�����쐬���܂��B�W�v�T�}���[�f�[�^��I�����Ă��������B", , "MCS 2020 - Print_spreadsheet"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        wb.Activate
        ws_mainmenu.Select
        End
    End If

'�E�U�C��ŕ\�����Ȃ��悤�ɂ��� - 2018/6/21
'    MsgBox "����ݒ������W�v�\Excel�t�@�C����I�����Ă��������B", , "MCS 2020 - Print_spreadsheet"

step00:
    spread_fn = Application.GetOpenFilename("�W�v�\Excel�t�@�C��,*.xlsx", , "�W�v�\Excel�t�@�C�����J��")
    If spread_fn = "False" Then
        ' �L�����Z���{�^���̏���
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf spread_fn = "" Then
        MsgBox "�W�v�\Excel�t�@�C����I�����Ă��������B", vbExclamation, "MCS 2020 - Print_spreadsheet"
        wb.Activate
        ws_mainmenu.Select
        GoTo step00
    ElseIf InStr(spread_fn, "_�W�v�\") = 0 Then
        MsgBox "�W�v�\Excel�t�@�C����I�����Ă��������B", vbExclamation, "MCS 2020 - Print_spreadsheet"
        wb.Activate
        ws_mainmenu.Select
        GoTo step00
    End If

    Workbooks.Open spread_fn
    ' �t���p�X����t�H���_���̎擾
    yen_pos = InStrRev(spread_fn, "\")
    spread_fd = Left(spread_fn, yen_pos - 1)
    
    spread_fd = spread_fd & "\����p"
    If Dir(spread_fd, vbDirectory) = "" Then
        MkDir spread_fd
    End If
    
    ' �t���p�X����t�@�C�����̎擾
    spread_fn = Dir(spread_fn)
    
    Set wb_spread = Workbooks(spread_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

    max_cnt = WorksheetFunction.CountA(ws_spread0.Columns(1)) - 1

'2018/06/19 - �ǋL ==========================
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - ����p�W�v�\�t�@�C���̍쐬"
    Form_Progress.Repaint
    progress_msg = "����p�W�v�\�t�@�C���̍쐬���L�����Z�����܂����B"
    Application.Visible = False
    AppActivate Form_Progress.Caption

    Form_Progress.Label1.Caption = "�����ݒ蒆"
    Form_Progress.Label2.Caption = "���΂炭���҂���������..."
    Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
    DoEvents
'============================================

    Call paradigm_procedure

'�W�v�\ �ڍ׈���ݒ�
'----N%
    '�\���܂�Ԃ�����
    ws_spread1.Activate
    If tn_tab = True Then
' MCODE�����̂Ƃ��ɓs���������̂łƂ肠�����R�����g�A�E�g
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    If ed_rw = st_rw Then
        MsgBox "C��ɕ\����񂪑��݂��Ȃ��ׁA�����𒆒f���܂��B"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP1/6 ����p�W�v�\�t�@�C���i�m���\�j�\���܂�Ԃ�������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        '�\�����ԏ�ɂ���J�e�S���̍s���擾
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        '�\���E�[�̗�ԍ����擾
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        '�\���������������A�J�e�S���[���l�d�b�d�ɒ���
        dmy_rw = st_rw + gt_rw
        Do While ed_cl > 16
            del_rw = Cells(dmy_rw, 7).End(xlDown).Row + 1
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = Cells(dmy_rw, 7).End(xlDown).Row
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")
    
    '�\���������� N%�\
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP2/6 ����p�W�v�\�t�@�C���i�m���\�j�\������������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        '�\������̊m�F
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 2
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        '�\�������L���̊m�F�i�\������25�ȏ�ɑ΂����s�j
        If ctz_cnt > 24 Then
            '�\���������̌v�Z
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            '�\������������
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 2), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 3, 1).Insert
            Next cpy
            '�J�e�S���[���l�d�b�d�ɒ���
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  '�P�i�S�́j�`�Q�R���ږڂ̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt * 2 - 1
                            del_st = del_st + 24 * 2
                            '�\�����x������ �@
                            '---------------------------------------------------------------------------
                            '�Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '�ȉ��͍����t�F�[�X������ɕK�v�ȏ���
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap
                    Case 2 To cpy_cnt - 1  '�Q�S���ږڂ���R�s�[���������̕\�̂ЂƂO�̍��ڂ܂�
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) * 2 - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 * 2 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt * 2 - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24 * 2, 4), Cells(del_ed + 24 * 2, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            '�\�����x������ �A
                            '---------------------------------------------------------------------------
                            '�P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 * 2 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24 * 2, 3).End(xlDown).Row
                        Next wrap
                    Case Is = cpy_cnt  '���������̕\�̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) * 2 - 1
                            '�\�����x������ �B
                            '---------------------------------------------------------------------------
                            '�Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
            
                            '�\�����x������ �C
                            '---------------------------------------------------------------------------
                            '�P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw * 2, 3).End(xlDown).Row
                        Next wrap
                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'----

'=====��������N�\
    ws_spread2.Activate
    If tn_tab = True Then
' MCODE�����̂Ƃ��ɓs���������̂łƂ肠�����R�����g�A�E�g
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    ' �J�e�S���s�̍�����12pt����24p�֕ύX
    st_rw = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 3).End(xlDown).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    Do
        For dmy_rw = st_rw To ed_rw
            Rows(dmy_rw).RowHeight = 24
        Next dmy_rw
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(Cells(ed_rw, 1).End(xlUp).Row, 3).End(xlDown).Row
        If ed_rw = 1 Then Exit Do
    Loop
    ' �\���܂�Ԃ�����
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    If ed_rw = st_rw Then
        MsgBox "C��ɕ\����񂪑��݂��Ȃ��ׁA�����𒆒f���܂��B"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP3/6 ����p�W�v�\�t�@�C���i�m�\�j�\���܂�Ԃ�������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        '�u�S�́v�J�e�S���̕\���ʒu�̐����ʒu��ύX
        If Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value = 0 Then
            Cells(Cells(st_rw, 3).End(xlDown).Row, 4).VerticalAlignment = xlTop
        End If
        ' �\�����ԏ�ɂ���J�e�S���̍s���擾
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        ' �\���E�[�̗�ԍ����擾
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        ' �\���������������A�J�e�S���[���l�d�b�d�ɒ���
        dmy_rw = st_rw + gt_rw
        del_rw = dmy_rw
        Do While ed_cl > 16
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    del_rw = del_rw + 1
                    Exit Do
                End If
            Loop
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = dmy_rw
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    Exit Do
                End If
            Loop
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")

    ' �\���������� N�\
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP4/6 ����p�W�v�\�t�@�C���i�m�\�j�\������������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        ' �\������̊m�F
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 1
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        ' �\�������L���̊m�F�i�\������25�ȏ�ɑ΂����s�j
        If ctz_cnt > 24 Then
            ' �\���������̌v�Z
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            ' �\������������
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 1), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 2, 1).Insert
            Next cpy
            ' �J�e�S���[��MECE�ɒ���
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  ' �P�i�S�́j�`�Q�R���ږڂ̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt - 1
                            del_st = del_st + 24
                            ' �\�����x������ �@
                            '---------------------------------------------------------------------------
                            ' �Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            ' �ȉ��͍����t�F�[�X������ɕK�v�ȏ���
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap

                    Case 2 To cpy_cnt - 1  ' �Q�S���ږڂ���R�s�[���������̕\�̂ЂƂO�̍��ڂ܂�
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24, 4), Cells(del_ed + 24, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ' �\�����x������ �A
                            '---------------------------------------------------------------------------
                            ' �P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24, 3).End(xlDown).Row
                        Next wrap

                    Case Is = cpy_cnt  ' ���������̕\�̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) - 1
                            ' �\�����x������ �B
                            '---------------------------------------------------------------------------
                            ' �Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With

                            ' �\�����x������ �C
                            '---------------------------------------------------------------------------
                            ' �P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw, 3).End(xlDown).Row
                        Next wrap

                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'=====�����܂�N�\

'=====��������%�\
    ws_spread3.Activate
    If tn_tab = True Then
' MCODE�����̂Ƃ��ɓs���������̂łƂ肠�����R�����g�A�E�g
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    ' �J�e�S���s�̍�����12pt����24p�֕ύX
    st_rw = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 3).End(xlDown).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    Do
        For dmy_rw = st_rw To ed_rw
            Rows(dmy_rw).RowHeight = 24
        Next dmy_rw
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(Cells(ed_rw, 1).End(xlUp).Row, 3).End(xlDown).Row
        If ed_rw = 1 Then Exit Do
    Loop
    ' �\���܂�Ԃ�����
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    If ed_rw = st_rw Then
        MsgBox "C��ɕ\����񂪑��݂��Ȃ��ׁA�����𒆒f���܂��B"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP5/6 ����p�W�v�\�t�@�C���i���\�j�\���܂�Ԃ�������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        '�u�S�́v�J�e�S���̕\���ʒu�̐����ʒu��ύX
        If Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value = 0 Then
            Cells(Cells(st_rw, 3).End(xlDown).Row, 4).VerticalAlignment = xlTop
        End If
        ' �\�����ԏ�ɂ���J�e�S���̍s���擾
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        ' �\���E�[�̗�ԍ����擾
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        ' �\���������������A�J�e�S���[���l�d�b�d�ɒ���
        dmy_rw = st_rw + gt_rw
        del_rw = dmy_rw
        Do While ed_cl > 16
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    del_rw = del_rw + 1
                    Exit Do
                End If
            Loop
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = dmy_rw
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    Exit Do
                End If
            Loop
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")

    ' �\���������� %�\
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP6/6 ����p�W�v�\�t�@�C���i���\�j�\������������" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        ' �\������̊m�F
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 1
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        ' �\�������L���̊m�F�i�\������25�ȏ�ɑ΂����s�j
        If ctz_cnt > 24 Then
            ' �\���������̌v�Z
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            ' �\������������
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 1), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 2, 1).Insert
            Next cpy
            ' �J�e�S���[��MECE�ɒ���
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  ' �P�i�S�́j�`�Q�R���ږڂ̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt - 1
                            del_st = del_st + 24
                            ' �\�����x������ �@
                            '---------------------------------------------------------------------------
                            ' �Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            ' �ȉ��͍����t�F�[�X������ɕK�v�ȏ���
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap

                    Case 2 To cpy_cnt - 1  ' �Q�S���ږڂ���R�s�[���������̕\�̂ЂƂO�̍��ڂ܂�
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24, 4), Cells(del_ed + 24, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ' �\�����x������ �A
                            '---------------------------------------------------------------------------
                            ' �P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24, 3).End(xlDown).Row
                        Next wrap

                    Case Is = cpy_cnt  ' ���������̕\�̏���
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) - 1
                            ' �\�����x������ �B
                            '---------------------------------------------------------------------------
                            ' �Q�S���ږڂ�������\�����x�����擾
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With

                            ' �\�����x������ �C
                            '---------------------------------------------------------------------------
                            ' �P���ږڂ�������\�����x�����m�F���A�󔒂Ȃ�\�����x�����㏑������
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw, 3).End(xlDown).Row
                        Next wrap

                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'=====�����܂�%�\

'---
'���Őݒ�
    ' N%�\�̉��Őݒ�
    ws_spread1.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    ' �ڎ��łւ̃y�[�W�ԍ��t�^�i�擪�j
    With ws_spread0.Cells(1, 6)
        .Value = "��"
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    s = 1
    not_found = "Not Found" & vbCrLf
    ws_spread0.Cells(2, 6).Value = s
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            ' �ڎ��łւ̃y�[�W�ԍ��t�^
            Set tab_no_fc = ws_spread0.Columns(2).Find(What:=ws_spread1.Cells(st_rw, 1).Value, lookat:=xlWhole)
            If ws_spread0.Cells(tab_no_fc.Row, 6).Value = "" Then
                s = s + 1
            End If
            If tab_no_fc Is Nothing Then
                not_found = not_found & ws_spread1.Cells(st_rw, 1).Value & vbCrLf
            ElseIf ws_spread1.Cells(st_rw, 1).Value <> "" Then
                ws_spread0.Cells(tab_no_fc.Row, 6).Value = s
                t = ws_spread0.Cells(tab_no_fc.Row, 6).Row
                Do
                    If ws_spread0.Cells(t, 6).Offset(-1, 0).Value <> "" Then
                        Exit Do
                    Else
                        ws_spread0.Cells(t, 6).Offset(-1, 0).Value = s - 1
                    End If
                    t = ws_spread0.Cells(t, 6).Offset(-1, 0).Row
                Loop
            End If
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 4).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                            ' ���ł��y�[�W�ԍ��ɔ��f
                            If dmy_rw > st_rw Then
                                s = s + 1
                            End If
                            '----
                        End If
                        dmy_rw = wrap
                    End If
                    ' �ŏI�y�[�W�̉��y�[�W����
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' N�\�̉��Őݒ�
    ws_spread2.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 1).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                        End If
                        dmy_rw = wrap
                    End If
                    ' �ŏI�y�[�W�̉��y�[�W����
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' %�\�̉��Őݒ�
    ws_spread3.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 1).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                        End If
                        dmy_rw = wrap
                    End If
                    '�ŏI�y�[�W�̉��y�[�W����
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' �y�[�W�ԍ��̏����ݒ� - 2018.9.20
    ws_spread0.Activate
    ws_spread0.Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""_ ;_ @_ "
    ws_spread0.Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' �\���̏����ݒ� - 2020.3.30
    ws_spread0.Activate
    ws_spread0.Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ShrinkToFit = True
    ws_spread0.Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' 2020.4.2 - �o�͂��P�y�[�W�����Ȃ��Ƃ��̕s��C��
    ws_spread0.Activate
    max_row = Cells(Rows.Count, 1).End(xlUp).Row
    For p_cnt = 2 To max_row
        If Cells(p_cnt, 6) <> "" Then
            p_temp = Cells(p_cnt, 6)
        Else
            Cells(p_cnt, 6) = p_temp
        End If
    Next p_cnt

'2018/06/19 - �ǋL ==========================
    Application.Visible = True
    Unload Form_Progress
'============================================

' �W�v�\����ݒ肱���܂�
'======
' �\���ݒ�
    Call cover_procedure

'======================================================================
' �o�c�e�o�͂�������
'======================================================================
    If tn_tab = True Then
        print_label = Replace(spread_fn, "_�W�v�\", "_�P���W�v�\")
    Else
        print_label = spread_fn
    End If
    wb_spread.SaveCopyAs spread_fd & "\" & "�y����pTEMP�z" & print_label
    For s = 1 To 3
        Select Case s
            Case 1
                Workbooks.Open Filename:=(spread_fd & "\" & "�y����pTEMP�z" & print_label)
                Set wb_print = Workbooks("�y����pTEMP�z" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(5).Delete
                wb_print.Worksheets(4).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\�i����p�j" & print_label)
                Application.DisplayAlerts = True
                Workbooks("�i����p�j" & print_label).Activate
                Call publish_procedure
                Workbooks("�y����pTEMP�z" & print_label).Close
                Workbooks("�i����p�j" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing

            Case 2
                Workbooks.Open Filename:=(spread_fd & "\" & "�y����pTEMP�z" & print_label)
                Set wb_print = Workbooks("�y����pTEMP�z" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(5).Delete
                wb_print.Worksheets(3).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\�i�����\�j" & print_label)
                Application.DisplayAlerts = True
                Workbooks("�i�����\�j" & print_label).Activate
                Call publish_procedure
                Workbooks("�y����pTEMP�z" & print_label).Close
                Workbooks("�i�����\�j" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing

            Case 3
                Workbooks.Open Filename:=(spread_fd & "\" & "�y����pTEMP�z" & print_label)
                Set wb_print = Workbooks("�y����pTEMP�z" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(4).Delete
                wb_print.Worksheets(3).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\�i�\����\�j" & print_label)
                Application.DisplayAlerts = True
                Workbooks("�i�\����\�j" & print_label).Activate
                Call publish_procedure
                Workbooks("�y����pTEMP�z" & print_label).Close
                Workbooks("�i�\����\�j" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing
            Case Else
        End Select
    Next s
    Application.DisplayAlerts = False
    wb_spread.Activate
    Kill spread_fd & "\" & "�y����pTEMP�z" & print_label
    wb_spread.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
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
    
' �V�X�e�����O�̏o�� - 2020.5.14
    ' 2020.6.3 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "26"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 26"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - ����p�W�v�\�t�@�C���̍쐬�F�Ώۃt�@�C���m" & spread_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "����p�W�v�\�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Print_spreadsheet"
End Sub

Private Sub paradigm_procedure()
' �ڎ��y�[�W�iws_spread0�j�̈���ݒ�
    Dim max_row As Long
    Dim i_cnt As Long
    Dim now_row As Long
    
    ws_spread0.Activate
    max_row = Cells(Rows.Count, 1).End(xlUp).Row

    ' �ڎ��s�̍�������
    For i_cnt = 2 To max_row
        now_row = Rows(i_cnt).RowHeight
        now_row = now_row / 12.75    ' �ڎ��P���ڂ̍s�����Z�o
        now_row = (now_row / 5) * 12.75
        Rows(i_cnt).RowHeight = Rows(i_cnt).RowHeight + now_row
    Next i_cnt
    
    If WorksheetFunction.CountA(Columns("D")) - 1 = 0 Then
        tn_tab = True
    Else
        tn_tab = False
    End If
    If Cells(1, 3).Value = "MCODE" Then
        Columns(3).Delete
    End If
    With Columns("A:B")
        .ColumnWidth = 4
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Columns("C:E")
        .ColumnWidth = 42
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    With Columns("F:H")
        .ColumnWidth = 6
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    Columns(6).Insert
    With Columns(6)
        .ColumnWidth = 4
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$A:$F"
        .PrintTitleRows = "$1:$1"               ' ����^�C�g���s�̎w��
        .PrintHeadings = False                  ' �s��ԍ����܂߂Ĉ�������Ȃ�
        .PrintGridlines = False                 ' �g������������Ȃ�
        .PrintComments = xlPrintNoComments      ' �Z���̃R�����g����������Ȃ�
        .PrintQuality = 600                     ' ����i����600dpi
        .CenterHorizontally = False             ' �Œ����i�����j�Ɉ�������Ȃ�
        .CenterVertically = False               ' �Œ����i�����j�Ɉ�������Ȃ�
        .Orientation = xlLandscape              ' �����������
        .Draft = False                          ' �ȈՈ�������Ȃ�
        .PaperSize = xlPaperA4                  ' �ŃT�C�Y��A4
        .Order = xlDownThenOver                 ' �Ŕԍ��̕t�ԋK�����ォ�牺
        .BlackAndWhite = False                  ' ������������Ȃ�
        .Zoom = 100                             ' ����{��
        .PrintErrors = xlPrintErrorsDisplayed   ' �G���[�\���̈���������܂܈��
        ' �ȉ���]���ݒ�
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(0)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
    End With
    Columns("G:I").Delete
    If tn_tab = True Then
        Columns("C").ColumnWidth = 9.88
        Columns("D:E").ColumnWidth = 58
        Columns("G").ColumnWidth = 4
    End If

    
    ' �ڎ��̍ŏI����
    Range("A1").Select
    Selection.CurrentRegion.Select
    With Selection.Font
        .Name = "���S�V�b�N"
        .Size = 8
    End With
    Rows("1:1").Select
    With Selection.Font
        .Name = "���S�V�b�N"
        .Size = 9
    End With
    Range("A1").Select

' �W�v�\�y�[�W�̐ݒ�
    ' N%
    ws_spread1.Activate
    Columns("C").ColumnWidth = 3
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' �s��ԍ����܂߂Ĉ�������Ȃ�
        .PrintGridlines = False                 ' �g������������Ȃ�
        .PrintComments = xlPrintNoComments      ' �Z���̃R�����g����������Ȃ�
        .PrintQuality = 600                     ' ����i����600dpi
        .CenterHorizontally = False             ' �Œ����i�����j�Ɉ�������Ȃ�
        .CenterVertically = False               ' �Œ����i�����j�Ɉ�������Ȃ�
        .Orientation = xlLandscape              ' �����������
        .Draft = False                          ' �ȈՈ�������Ȃ�
        .PaperSize = xlPaperA4                  ' �ŃT�C�Y��A4
        .FirstPageNumber = 1                    ' �擪�Ŕԍ���1
        .Order = xlDownThenOver                 ' �Ŕԍ��̕t�ԋK�����ォ�牺
        .BlackAndWhite = False                  ' ������������Ȃ�
        .Zoom = 80                              ' ����{��
        .PrintErrors = xlPrintErrorsDisplayed   ' �G���[�\���̈���������܂܈��
        ' �ȉ��A�]���ݒ�
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        '�ȉ��A�t�b�^�[�ݒ�
        .RightFooter = "&""Century""&9&P"
    End With
    ' N
    ws_spread2.Activate
    With Columns("C")
        .ColumnWidth = 3
        .VerticalAlignment = xlTop
    End With
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
        .VerticalAlignment = xlTop
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' �s��ԍ����܂߂Ĉ�������Ȃ�
        .PrintGridlines = False                 ' �g������������Ȃ�
        .PrintComments = xlPrintNoComments      ' �Z���̃R�����g����������Ȃ�
        .PrintQuality = 600                     ' ����i����600dpi
        .CenterHorizontally = False             ' �Œ����i�����j�Ɉ�������Ȃ�
        .CenterVertically = False               ' �Œ����i�����j�Ɉ�������Ȃ�
        .Orientation = xlLandscape              ' �����������
        .Draft = False                          ' �ȈՈ�������Ȃ�
        .PaperSize = xlPaperA4                  ' �ŃT�C�Y��A4
        .FirstPageNumber = 1                    ' �擪�Ŕԍ���1
        .Order = xlDownThenOver                 ' �Ŕԍ��̕t�ԋK�����ォ�牺
        .BlackAndWhite = False                  ' ������������Ȃ�
        .Zoom = 80                              ' ����{��
        .PrintErrors = xlPrintErrorsDisplayed   ' �G���[�\���̈���������܂܈��
        ' �ȉ��A�]���ݒ�
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        ' �ȉ��A�t�b�^�[�ݒ�
        .RightFooter = "&""Century""&9&P"
    End With
    ' %
    ws_spread3.Activate
    With Columns("C")
        .ColumnWidth = 3
        .VerticalAlignment = xlTop
    End With
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
        .VerticalAlignment = xlTop
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' �s��ԍ����܂߂Ĉ�������Ȃ�
        .PrintGridlines = False                 ' �g������������Ȃ�
        .PrintComments = xlPrintNoComments      ' �Z���̃R�����g����������Ȃ�
        .PrintQuality = 600                     ' ����i����600dpi
        .CenterHorizontally = False             ' �Œ����i�����j�Ɉ�������Ȃ�
        .CenterVertically = False               ' �Œ����i�����j�Ɉ�������Ȃ�
        .Orientation = xlLandscape              ' �����������
        .Draft = False                          ' �ȈՈ�������Ȃ�
        .PaperSize = xlPaperA4                  ' �ŃT�C�Y��A4
        .FirstPageNumber = 1                    ' �擪�Ŕԍ���1
        .Order = xlDownThenOver                 ' �Ŕԍ��̕t�ԋK�����ォ�牺
        .BlackAndWhite = False                  ' ������������Ȃ�
        .Zoom = 80                              ' ����{��
        .PrintErrors = xlPrintErrorsDisplayed   ' �G���[�\���̈���������܂܈��
        ' �ȉ��A�]���ݒ�
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        ' �ȉ��A�t�b�^�[�ݒ�
        .RightFooter = "&""Century""&9&P"
    End With
End Sub
 
Private Sub publish_procedure()
    Dim PathPdf As String
    PathPdf = spread_fd & "\" & Replace(ActiveWorkbook.Name, ".xlsx", ".pdf")
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PathPdf, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

Private Sub cover_procedure()
' �\���ݒ�
    Dim objFSO As Object, wb_obj As Object
    Dim wb_cover As Workbook, wb_crs As Workbook, wb_rd As Workbook
    Dim ws_cover As Worksheet
    Dim cover_fd As String, crs_fd As String, crs_fn As String, rd_fd As String, rd_fn As String
    Dim s_rw As Long, s_cnt As Long
    Dim tBox_Ctrl As Shape
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    cover_fd = file_path & "\5_INI"
    For Each wb_obj In objFSO.getfolder(cover_fd).Files
        If Right(wb_obj.Name, 8) = "cov.xlsx" Then
            Set wb_cover = Workbooks.Open(cover_fd & "\" & wb_obj.Name)
            Set ws_cover = wb_cover.Worksheets(1)
            ws_cover.Name = "�\��"
            ws_cover.Move before:=ws_spread0
            Application.DisplayAlerts = False
            wb_cover.Activate
            wb_cover.Close
            Application.DisplayAlerts = True

            Set ws_cover = wb_spread.Worksheets(1)
            
            For Each tBox_Ctrl In ws_cover.Shapes
                If tBox_Ctrl.Type = 17 Then
                    Select Case tBox_Ctrl.TextFrame.Characters.Text
                        Case "�^�C�g��1"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 3 Then
                                If tn_tab = True And ws_mainmenu.Cells(3, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value & " �P���W�v�\"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(3, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value & " �W�v�\"
                                End If
                            ElseIf ws_mainmenu.Cells(6, 32).End(xlUp).Row > 3 Then
                                tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "�^�C�g��1" Then
                                tBox_Ctrl.Delete
                            End If
                        Case "�^�C�g��2"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 4 Then
                                If tn_tab = True And ws_mainmenu.Cells(4, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value & " �P���W�v�\"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(4, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value & " �W�v�\"
                                End If
                            ElseIf ws_mainmenu.Cells(6, 32).End(xlUp).Row > 4 Then
                                tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "�^�C�g��2" Then
                                tBox_Ctrl.Delete
                            End If
                        Case "�^�C�g��3"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 5 Then
                                If tn_tab = True And ws_mainmenu.Cells(5, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(5, 32).Value & " �P���W�v�\"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(5, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(5, 32).Value & " �W�v�\"
                                End If
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "�^�C�g��3" Then
                                tBox_Ctrl.Delete
                            End If
' �\���Ɍ�������̉�����a������̂ŁA������ƕۗ��c - 2018/6/22
'                        Case "�W�v�Ώی����F"
'                            crs_fd = file_path & "\3_FD"
'                            crs_fn = Replace(spread_fn, "_�W�v�\", "")
'                            Set wb_crs = Workbooks.Open(crs_fd & "\" & crs_fn)
'                            If WorksheetFunction.CountA(wb_crs.Worksheets(1).Range("Q:R")) = 5 Then
'                                rd_fn = wb_crs.Worksheets(1).Cells(2, 4).Value
'                                rd_fd = file_path & "\1_DATA"
'                                Set wb_rd = Workbooks.Open(rd_fd & "\" & rd_fn)
'                                s_rw = 6
'                                Do
'                                    If Cells(s_rw, 1).Offset(1, 0).Value <> "" And Cells(s_rw, 1).Value <> Cells(s_rw, 1).Offset(1, 0).Value Then
'                                        s_cnt = s_cnt + 1
'                                    ElseIf Cells(s_rw, 1).Offset(1, 0).Value = "" Then
'                                        Exit Do
'                                    End If
'                                    s_rw = s_rw + 1
'                                Loop
'                                tBox_Ctrl.TextFrame.Characters.Text = tBox_Ctrl.TextFrame.Characters.Text & s_cnt & " ��"
'                            End If
'                            wb_crs.Close
'                            wb_rd.Close
                        Case Else
                    End Select
                    ActiveWindow.View = xlNormalView
                    ActiveWindow.DisplayGridlines = False
                    ActiveWindow.Zoom = 80
                End If
            Next
            With ws_cover.Range("AB49")
                With .Font
                    .Name = "Arial"
                    .Color = RGB(255, 255, 255)
                End With
                .Value = "ACROSS Multiple Cross-tabulation System in" & Str(Year(Now)) & Space(1)
                .HorizontalAlignment = xlRight
                .FontSize = 8
            End With
        End If
    Next wb_obj
    
    Set wb_obj = Nothing
    Set ws_cover = Nothing
    Set wb_cover = Nothing
    Set wb_crs = Nothing
End Sub

Private Sub CoverMark_procedure()
    Dim ws_cover As Worksheet
    Dim tBox_Ctrl As Shape
    Set ws_cover = wb_print.Worksheets(1)
    For Each tBox_Ctrl In ws_cover.Shapes
        If tBox_Ctrl.Type = 17 And tBox_Ctrl.TextFrame.Characters.Text = "�o�̓^�C�v" Then
            Select Case wb_print.Worksheets(3).Name
                Case "�m���\"
                    tBox_Ctrl.Delete
                Case "�m�\"
                    tBox_Ctrl.TextFrame.Characters.Text = "�����\"
                Case "���\"
                    tBox_Ctrl.TextFrame.Characters.Text = "�\����\"
                Case Else
            End Select
        End If
    Next
    Set ws_cover = Nothing
End Sub

