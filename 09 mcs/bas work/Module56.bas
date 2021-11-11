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
'2018/06/01 - �ǋL ==========================
    Dim r_code As Integer
    Dim spd_tab() As String
    Dim spd_file As String
    Dim spd_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long
'--------------------------------------------------------------------------------------------------'
'�@���K�V�[�ŏW�v�\CSV�t�@�C���̍쐬 �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.11.15�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "���K�V�[�ŏW�v�\CSV�t�@�C���̍쐬��..."
    
    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    ' SUM�t�H���_����*_�W�v�\.xlsx�`���̃t�@�C�������J�E���g
    spd_cnt = 0
    spd_file = Dir(file_path & "\SUM\*_�W�v�\.xlsx")
    Do Until spd_file = ""
        DoEvents
        spd_cnt = spd_cnt + 1
        spd_file = Dir()
    Loop
    
    ' SUM�t�H���_����*_�W�v�\.xlsx�`���̃t�@�C������z��ɃZ�b�g
    ReDim spd_tab(spd_cnt)
    spd_file = Dir(file_path & "\SUM\*_�W�v�\.xlsx")
    For fn_cnt = 1 To spd_cnt
        DoEvents
        spd_tab(fn_cnt) = spd_file
        spd_file = Dir()
    Next fn_cnt
    fn_cnt = spd_cnt
  
    rc = MsgBox("�W�v�\Excel�t�@�C������A���K�V�[�ŏW�v�\CSV�t�@�C�����쐬���܂��B" & vbCrLf & "�쐬�ΏۂƂȂ�W�v�\Excel�t�@�C���͂���܂����B" _
      & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "�W�v�\CSV�t�@�C�����쐬���邽�߂ɕK�v�ȏW�v�\Excel�t�@�C�����Ȃ��ꍇ�́u�������v��I�����Ă��������B", vbYesNoCancel + vbQuestion, "�W�v�\Excel�t�@�C���쐬�̊m�F")
    If rc = vbNo Then
        MsgBox "�W�v�\Excel�t�@�C�����쐬���܂��B�W�v�T�}���[�f�[�^��I�����Ă��������B"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        Call Finishing_Mcs2017
        End
    End If

' ���K�V�[�ŏW�v�\CSV�t�@�C�������쐬�����i������A�܂��͂P�񏈗����璅�肵�Ă��܂��j
    If spd_cnt > 0 Then
        r_code = MsgBox("SUM�t�H���_���ɂ���" & fn_cnt & "�̏W�v�\Excel�t�@�C������A" & vbCrLf & "�ꊇ���ă��K�V�[�ŏW�v�\CSV�t�@�C�����쐬���܂����B" _
         & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "SUM�t�H���_���́m*_�W�v�\.xlsx�`���n�̃t�@�C������" & vbCrLf & "�\�����Ă��܂��B" _
         & vbCrLf & "�u�͂��v�@�� �W�v�\Excel�t�@�C�����ꊇ����" & vbCrLf & "�u�������v�� �W�v�\Excel�t�@�C����I�����Ă��珈��", _
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

                ' �t�@�C��������g���q�ȊO�̎擾
                csv_fn = Left(spread_fn, InStr(spread_fn, "_�W�v�\") - 1)

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

                Sheets("�ڎ�").Select
                Call cells_format
                Call index_format
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_�ڎ�.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("�m���\").Select
                Call cells_format
                Call legacy_format1
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP�\.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("�m�\").Select
                Call cells_format
                Call legacy_format2
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N�\.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                
                Sheets("���\").Select
                Call cells_format
                Call legacy_format3
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_P�\.csv", _
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

' �V�X�e�����O�̏o��
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
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�\CSV�t�@�C���̍쐬�F�Ώۃt�@�C���mSUM�t�H���_����" & spd_cnt - 1 & "�̏W�v�\Excel�t�@�C���n"
            Close #1
            Call Finishing_Mcs2017
            MsgBox spd_cnt - 1 & "�̏W�v�\CSV�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Csv_spreadsheet"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If

' ���K�V�[�ŏW�v�\CSV�t�@�C���P��쐬�����i�P�񏈗����璅��A�s�Ӎ쐬���j
step00:
    wb.Activate
    ws_mainmenu.Select
' ���邳���̂ŁA���L���b�Z�[�W���R�����g�A�E�g
'    MsgBox "���K�V�[�ŏW�v�\CSV�t�@�C�����쐬����W�v�\Excel�t�@�C����I�����Ă��������B"
    spread_fn = Application.GetOpenFilename("�W�v�\Excel�t�@�C��,*.xlsx", , "�W�v�\Excel�t�@�C�����J��")
    If spread_fn = "False" Then
        ' �L�����Z���{�^���̏���
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf spread_fn = "" Then
        MsgBox "�W�v�\Excel�t�@�C����I�����Ă��������B", vbExclamation, "MCS 2020 - Print_spreadsheet"
        GoTo step00
    ElseIf InStr(spread_fn, "_�W�v�\") = 0 Then
        MsgBox "�W�v�\Excel�t�@�C����I�����Ă��������B", vbExclamation, "MCS 2020 - Print_spreadsheet"
        GoTo step00
    End If

    Open spread_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(spread_fn).Close
    Else
        Workbooks.Open spread_fn
    End If
    
    ' �t���p�X����t�H���_���̎擾
    yen_pos = InStrRev(spread_fn, "\")
    spread_fd = Left(spread_fn, yen_pos - 1)
    
    ' �t���p�X����t�@�C�����̎擾
    spread_fn = Dir(spread_fn)
    
    ' �t�@�C��������g���q�ȊO�̎擾
    csv_fn = Left(spread_fn, InStr(spread_fn, "_�W�v�\") - 1)
    
    Set wb_spread = Workbooks(spread_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

' ���K�V�[�ŏW�v�\CSV�t�@�C���쐬��������
    
    Application.DisplayAlerts = False
    wb_spread.Activate

    csv_fd = spread_fd & "\CSV\"
    If Dir(csv_fd, vbDirectory) = "" Then
        MkDir csv_fd
    End If
    
    Sheets("�ڎ�").Select
    Call cells_format
    Call index_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_�ڎ�.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("�m���\").Select
    Call cells_format
    Call legacy_format1
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP�\.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("�m�\").Select
    Call cells_format
    Call legacy_format2
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N�\.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("���\").Select
    Call cells_format
    Call legacy_format3
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_P�\.csv", _
        FileFormat:=xlCSV, CreateBackup:=False

    ActiveWorkbook.Close
    Application.DisplayAlerts = True

' �W�v�\CSV�t�@�C���쐬�����܂�
    
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
    
' �V�X�e�����O�̏o��
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�\CSV�t�@�C���̍쐬�F�Ώۃt�@�C���m" & spread_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "���K�V�[�ŏW�v�\CSV�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Csv_spreadsheet"
End Sub

Private Sub cells_format()
    Cells.Select
    With Selection
        .ClearFormats
    End With
    Range("A1").Select
End Sub

Private Sub index_format()
' �ڎ��t�H�[�}�b�g����
    Cells.Replace What:="" & Chr(10) & "", Replacement:="��", lookat:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Private Sub legacy_format1()
' �m���\�t�H�[�}�b�g����
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
    
    ' B��̍폜
    Columns("B").Delete
    
    ' �V�[�g�̍ŏI�s�擾
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' �W�v�\�P�\������̊J�n�s�擾
            bgn_row = ActiveCell.Row
            ' �W�v�\�P�\������̊J�n��擾
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            '�Z���N�g�̗L���m�F�ƃZ���N�g�R�����g�̎擾
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "�y" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "�z")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' �W�v�\�P�\������̍ŏI�s�擾
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' �������̍s��擾
            For k_cnt = bgn_row To fin_row
                If ws_spread1.Cells(k_cnt, 5) = "����" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' �W�v�\�P�\������̍ŏI��擾
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' �\��R�����g�̃Z�b�g
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "�\��"
            Cells(ken_row + 1, ken_col - 1) = "���v"
            Cells(ken_row, ken_col - 3).Select
            
            ' �Z���N�g�R�����g�̃Z�b�g
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "�W�v����"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "�F", "�c")
                Next p_cnt
            End If
            
            ' �\�����ڔԍ��̏���
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' �\���́u���v�v�u�W���΍��v�̏���
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "�W���΍�" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "���v" Then
                    Cells(ken_row, p_cnt) = "�������v"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' �����Ɠ��v�ʁi�������ځj�̒��� - 2019.12.12
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "����" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "����" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�ŏ��l" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "��P�l����" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�����l" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "��R�l����" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�ő�l" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�ŕp�l" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�W���΍�" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
                If Cells(ken_row, p_cnt) = "�������v" Then
                    For f_cnt = ken_row + (s_cnt - 1) + 1 To fin_row + (s_cnt - 1)
                        If Cells(f_cnt, p_cnt) = "" Then
                            Cells(f_cnt, p_cnt) = Cells(f_cnt - 1, p_cnt)
                        End If
                    Next f_cnt
                End If
            Next p_cnt
            
            ' �W�v�\�ԍ��̏����ƕs�v�ȍs�̍폜����
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
            
            ' ���̏W�v�\�̎n�_���Z���N�g
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C��̍폜
    Columns("C").Delete

End Sub

Private Sub legacy_format2()
' �m�\�t�H�[�}�b�g����
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
    
    ' B��̍폜
    Columns("B").Delete
    
    ' �V�[�g�̍ŏI�s�擾
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' �W�v�\�P�\������̊J�n�s�擾
            bgn_row = ActiveCell.Row
            ' �W�v�\�P�\������̊J�n��擾
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            '�Z���N�g�̗L���m�F�ƃZ���N�g�R�����g�̎擾
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "�y" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "�z")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' �W�v�\�P�\������̍ŏI�s�擾
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' �������̍s��擾
            For k_cnt = bgn_row To fin_row
                If ws_spread2.Cells(k_cnt, 5) = "����" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' �W�v�\�P�\������̍ŏI��擾
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' �\��R�����g�̃Z�b�g
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "�\��"
            Cells(ken_row + 1, ken_col - 1) = "���v"
            Cells(ken_row, ken_col - 3).Select
            
            ' �Z���N�g�R�����g�̃Z�b�g
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "�W�v����"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "�F", "�c")
                Next p_cnt
            End If
            
            ' �\�����ڔԍ��̏���
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' �\���́u���v�v�u�W���΍��v�̏���
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "�W���΍�" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "���v" Then
                    Cells(ken_row, p_cnt) = "�������v"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' �W�v�\�ԍ��̏����ƕs�v�ȍs�̍폜����
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
            
            ' ���̏W�v�\�̎n�_���Z���N�g
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C��̍폜
    Columns("C").Delete

End Sub

Private Sub legacy_format3()
' ���\�t�H�[�}�b�g����
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
    
    ' B��̍폜
    Columns("B").Delete
    
    ' �V�[�g�̍ŏI�s�擾
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    max_col = ActiveCell.Column
    
    Range("A1").Select
    
    wk_row = 1
    del_row = 0
    For r_cnt = 1 To max_row
        If wk_row < max_row Then
            Cells(wk_row, 1).Select
        
            ' �W�v�\�P�\������̊J�n�s�擾
            bgn_row = ActiveCell.Row
            ' �W�v�\�P�\������̊J�n��擾
            bgn_col = ActiveCell.Column
    
            Selection.End(xlDown).Select
    
            '�Z���N�g�̗L���m�F�ƃZ���N�g�R�����g�̎擾
            sel_flag = 0
            For s_cnt = 1 To 7
                If Mid(Cells(bgn_row + s_cnt, 3), 1, 1) = "�y" Then
                    sel_flag = 1
                    findNo = InStr(Cells(bgn_row + 1, 3), "�z")
                    sel_cm(s_cnt) = Right(Cells(bgn_row + s_cnt, 3), Len(Cells(bgn_row + s_cnt, 3)) - findNo)
                Else
                    Exit For
                End If
            Next s_cnt
            
            ' �W�v�\�P�\������̍ŏI�s�擾
            If ActiveCell.Row = Rows.Count Then
                fin_row = max_row
            Else
                fin_row = ActiveCell.Row - 2
            End If
            
            ' �������̍s��擾
            For k_cnt = bgn_row To fin_row
                If ws_spread3.Cells(k_cnt, 5) = "����" Then
                    ken_row = k_cnt
                    Exit For
                End If
            Next k_cnt
            Cells(ken_row, 5).Select
            ken_col = ActiveCell.Column
            
            ' �W�v�\�P�\������̍ŏI��擾
            Cells(ken_row, ken_col).Select
            fin_col = Cells(ken_row, Columns.Count).End(xlToLeft).Column
            
            ' �\��R�����g�̃Z�b�g
            Cells(ken_row, ken_col - 1) = Cells(bgn_row, bgn_col + 2)
            Cells(ken_row, ken_col - 3) = "�\��"
            Cells(ken_row + 1, ken_col - 1) = "���v"
            Cells(ken_row, ken_col - 3).Select
            
            ' �Z���N�g�R�����g�̃Z�b�g
            If sel_flag = 1 Then
                Rows(ActiveCell.Row + 1 & ":" & ActiveCell.Row + s_cnt - 1).Insert
                For p_cnt = 1 To s_cnt - 1
                    Cells(ken_row + p_cnt, 2) = "�W�v����"
                    Cells(ken_row + p_cnt, 4) = sel_cm(p_cnt)
                    Cells(ken_row + p_cnt, 4) = Replace(Cells(ken_row + p_cnt, 4), "�F", "�c")
                Next p_cnt
            End If
            
            ' �\�����ڔԍ��̏���
            Cells(ken_row, ken_col - 3).Select
            For p_cnt = ken_row To fin_row + (s_cnt - 1)
                If Cells(p_cnt, 2) = "" Then
                    Cells(p_cnt, 2) = Cells(p_cnt - 1, 2)
                End If
            Next p_cnt
            
            ' �\���́u���v�v�u�W���΍��v�̏���
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "�W���΍�" Then
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            Cells(ken_row, ken_col).Select
            For p_cnt = ken_col To fin_col
                If Cells(ken_row, p_cnt) = "���v" Then
                    Cells(ken_row, p_cnt) = "�������v"
                    Range(Cells(ken_row, p_cnt), Cells(fin_row, p_cnt)).Cut
                    Cells(ken_row, fin_col + 1).Insert
                End If
            Next p_cnt
            
            ' �W�v�\�ԍ��̏����ƕs�v�ȍs�̍폜����
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
            
            ' ���̏W�v�\�̎n�_���Z���N�g
            Cells(fin_row + (s_cnt - 1) - aj_row + 2, 1).Select
            wk_row = ActiveCell.Row
            max_row = max_row + (s_cnt - 1) - aj_row
        Else
            Exit For
        End If
    Next r_cnt

    ' C��̍폜
    Columns("C").Delete

End Sub


