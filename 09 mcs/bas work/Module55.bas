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
'2018/06/01 - �ǋL ==========================
    Dim r_code As Integer
    Dim spd_tab() As String
    Dim spd_file As String
    Dim spd_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long
'--------------------------------------------------------------------------------------------------'
'�@�W�v�\CSV�t�@�C���̍쐬 �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.04.27�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "�W�v�\CSV�t�@�C���̍쐬��..."
    
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
  
    rc = MsgBox("�W�v�\Excel�t�@�C������A�W�v�\CSV�t�@�C�����쐬���܂��B" & vbCrLf & "�쐬�ΏۂƂȂ�W�v�\Excel�t�@�C���͂���܂����B" _
      & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "�W�v�\CSV�t�@�C�����쐬���邽�߂ɕK�v�ȏW�v�\Excel�t�@�C�����Ȃ��ꍇ�́u�������v��I�����Ă��������B", vbYesNoCancel + vbQuestion, "�W�v�\Excel�t�@�C���쐬�̊m�F")
    If rc = vbNo Then
        MsgBox "�W�v�\Excel�t�@�C�����쐬���܂��B�W�v�T�}���[�f�[�^��I�����Ă��������B"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        Call Finishing_Mcs2017
        End
    End If

' �W�v�\CSV�t�@�C�������쐬����
    If spd_cnt > 0 Then
        r_code = MsgBox("SUM�t�H���_���ɂ���" & fn_cnt & "�̏W�v�\Excel�t�@�C������A" & vbCrLf & "�ꊇ���ďW�v�\CSV�t�@�C�����쐬���܂����B" _
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
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_�ڎ�.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("�m���\").Select
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP�\.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("�m�\").Select
                ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N�\.csv", _
                    FileFormat:=xlCSV, CreateBackup:=False
                Sheets("���\").Select
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

' �W�v�\CSV�t�@�C���P��쐬����
step00:
    wb.Activate
    ws_mainmenu.Select
    MsgBox "�W�v�\CSV�t�@�C�����쐬����W�v�\Excel�t�@�C����I�����Ă��������B"
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

' �W�v�\CSV�t�@�C���쐬��������
    
    Application.DisplayAlerts = False
    wb_spread.Activate

    csv_fd = spread_fd & "\CSV\"
    If Dir(csv_fd, vbDirectory) = "" Then
        MkDir csv_fd
    End If
    
    Sheets("�ڎ�").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_�ڎ�.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("�m���\").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_NP�\.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("�m�\").Select
    Call cells_format
    ActiveWorkbook.SaveAs Filename:=csv_fd & csv_fn & "_N�\.csv", _
        FileFormat:=xlCSV, CreateBackup:=False
    
    Sheets("���\").Select
    Call cells_format
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
    ' 2020.6.3 - �ǉ�
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�\CSV�t�@�C���̍쐬�F�Ώۃt�@�C���m" & spread_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�W�v�\CSV�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Csv_spreadsheet"
End Sub

Private Sub cells_format()
    Cells.Select
    With Selection
        .ClearFormats
    End With
    Range("A1").Select
End Sub
