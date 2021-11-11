Attribute VB_Name = "Module53"
Option Explicit
    Dim mcs_ini(10) As String
    Dim ini_cnt As Integer
    Dim wb_spread As Workbook
    Dim ws_spread0 As Worksheet
    Dim ws_spread1 As Worksheet
    Dim ws_spread2 As Worksheet
    Dim ws_spread3 As Worksheet
    Dim summary_fd As String
    Dim summary_fn As String
    Dim spread_fn As String
    Dim hyo_cnt As Long
    Dim max_row As Long
    Dim np_max_row As Long
    Dim face_flag As Integer
    Dim i_cnt As Long

Sub Spreadsheet_Creation()
    Dim waitTime As Variant
    Dim yen_pos As Long
'2018/05/28 - �ǋL ==========================
    Dim r_code As Integer
    Dim sum_tab() As String
    Dim sum_file As String
    Dim sum_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long
'2019/12/10 - �ǋL ==========================
    Dim wgt_row As Long
'--------------------------------------------------------------------------------------------------'
'�@�W�v�\Excel�t�@�C���̍쐬 �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.05.24�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    
    ' �ݒ�t�@�C������̏��擾
    ' (1)�p�X�A(2)���{��t�H���g�A(3)���{��t�H���g�T�C�Y�A(4)�p�����t�H���g�A(5)�p�����t�H���g�T�C�Y�A(6)�S�̗��J���[�A(7)�r���J���[
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
        Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
         "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Input As #1
        ini_cnt = 1
        Do Until EOF(1)
            DoEvents
            Line Input #1, mcs_ini(ini_cnt)
            Select Case ini_cnt
            Case 2
                If Mid(mcs_ini(ini_cnt), 1, 7) <> "J-FONT=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mFONT�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 3
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "J-FONT-SIZE=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mFONT-SIZE�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 4
                If Mid(mcs_ini(ini_cnt), 1, 7) <> "E-FONT=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mFONT�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 5
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "E-FONT-SIZE=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mFONT-SIZE�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 6
                If Mid(mcs_ini(ini_cnt), 1, 12) <> "TOTAL-COLOR=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mTOTAL-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 16, 1) <> "," Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mTOTAL-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 20, 1) <> "," Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mTOTAL-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case 7
                If Mid(mcs_ini(ini_cnt), 1, 13) <> "BORDER-COLOR=" Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mBORDER-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 17, 1) <> "," Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mBORDER-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
                If Mid(mcs_ini(ini_cnt), 21, 1) <> "," Then
                    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n�́mBORDER-COLOR�n�ݒ���m�F���Ă��������B" _
                     & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
                    Call Finishing_Mcs2017
                    End
                End If
            Case Else
            End Select
            ini_cnt = ini_cnt + 1
        Loop
        Close #1
    Else
        MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n��������܂���B" _
         & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        Call Finishing_Mcs2017
        End
    End If
    
    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    ' SUM�t�H���_����*_sum.xlsx�`���̃t�@�C�������J�E���g
    sum_cnt = 0
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    Do Until sum_file = ""
        DoEvents
        sum_cnt = sum_cnt + 1
        sum_file = Dir()
    Loop
    
    ' SUM�t�H���_����*_sum.xlsx�`���̃t�@�C������z��ɃZ�b�g
    ReDim sum_tab(sum_cnt)
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    For fn_cnt = 1 To sum_cnt
        DoEvents
        sum_tab(fn_cnt) = sum_file
        sum_file = Dir()
    Next fn_cnt
    fn_cnt = sum_cnt

' �W�v�\Excel�t�@�C�������쐬����
    If sum_cnt > 0 Then
        r_code = MsgBox("SUM�t�H���_���ɂ���" & fn_cnt & "�̏W�v�T�}���[�f�[�^����A" & vbCrLf & "�ꊇ���ďW�v�\Excel�t�@�C�����쐬���܂����B" _
         & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "SUM�t�H���_���́m*_sum.xlsx�`���n�̃t�@�C������" & vbCrLf & "�\�����Ă��܂��B" _
         & vbCrLf & "�u�͂��v�@�� �W�v�T�}���[�f�[�^���ꊇ����" & vbCrLf & "�u�������v�� �W�v�T�}���[�t�@�C����I�����Ă��珈��", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Spreadsheet_Creation")
        If r_code = vbYes Then
            sum_cnt = 1
            For n_cnt = 1 To fn_cnt
                DoEvents
                wb.Activate
                ws_mainmenu.Select
                summary_fn = sum_tab(n_cnt)
                
                Open file_path & "\SUM\" & summary_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(summary_fn).Close
                Else
                    Workbooks.Open summary_fn
                End If
                
                Set wb_spread = Workbooks(summary_fn)
                Set ws_spread0 = wb_spread.Worksheets(1)
                Set ws_spread1 = wb_spread.Worksheets(2)
                Set ws_spread2 = wb_spread.Worksheets(3)
                Set ws_spread3 = wb_spread.Worksheets(4)
                
                wb_spread.Activate
                Sheets(Array("�m���\", "�m�\", "���\")).Select
                Cells.Select
                
                ' �W�v�\�V�[�g�S�̂̃X�^�C���ݒ�
                ActiveWindow.DisplayGridlines = False
                With Selection.Font
                    .Name = Mid(mcs_ini(2), 8)      ' ���{��t�H���g���Z�b�g
                    .Size = Mid(mcs_ini(3), 13)     ' ���{��t�H���g�T�C�Y���Z�b�g
                End With
                
                ' �W�v�\�V�[�g�S�̂̕\���̃X�^�C���ݒ�
                Columns(1).Select
                With Selection
                    .ColumnWidth = 4.88
                End With
                
                ' �W�v�\�V�[�g�S�̂̕\���J�e�S���[�ԍ��̃X�^�C���ݒ�
                Columns(3).Select
                With Selection
                    .HorizontalAlignment = xlRight
                    .ColumnWidth = 4.25
                    .Font.Color = RGB(0, 112, 192)
                    .Font.Italic = True
                    .Font.Size = 7
                End With
                
                Range("A1").Select
                ws_spread0.Select
                ws_spread0.PageSetup.Orientation = xlLandscape
                hyo_cnt = ws_spread0.Cells(Rows.Count, setup_col).End(xlUp).Row - 1
                
                face_flag = 0       ' �N���X�W�v�\�̗L������p�t���O
                Application.ScreenUpdating = False
                Load Form_Progress
                Form_Progress.StartUpPosition = 1
                Form_Progress.Show vbModeless
                Form_Progress.Caption = "MCS 2020 - �W�v�\Excel�t�@�C���̍쐬"
                Form_Progress.Repaint
                progress_msg = "�W�v�\Excel�t�@�C���̍쐬���L�����Z�����܂����B"
                Application.Visible = False
                AppActivate Form_Progress.Caption
                
                ' �m���\�̏���
                ws_spread1.Select
                ws_spread1.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP1/5 �W�v�\Excel�t�@�C���i�m���\�j�쐬��" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
                    Call int_spreadsheet
                Next i_cnt
    
                ws_spread1.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "���S�V�b�N"
                    .Size = 8
                End With
                ws_spread1.Columns(2).Hidden = True     ' MCODE��̔�\��
                Range("A1").Select
                
                ' �m�\�̏���
                ws_spread2.Select
                ws_spread2.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
'---------------------------------
                np_max_row = ActiveCell.Row
'---------------------------------
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP2/5 �W�v�\Excel�t�@�C���i�m�\�j�쐬��" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
                    Call ken_spreadsheet
                Next i_cnt
                
                ws_spread2.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "���S�V�b�N"
                    .Size = 8
                End With
                ws_spread2.Columns(2).Hidden = True     ' MCODE��̔�\��
                Range("A1").Select
                
                ' ���\�̏���
                ws_spread3.Select
                ws_spread3.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
'---------------------------------
                np_max_row = ActiveCell.Row
'---------------------------------
                
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP3/5 �W�v�\Excel�t�@�C���i���\�j�쐬��" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
                    Call per_spreadsheet
                Next i_cnt
                
                ws_spread3.Columns("B:B").Select
                Selection.ClearFormats
                With Selection.Font
                    .Name = "���S�V�b�N"
                    .Size = 8
                End With
                ws_spread3.Columns(2).Hidden = True     ' MCODE��̔�\��
                Range("A1").Select
                
                Form_Progress.Label1.Caption = "100%"
                DoEvents
                Form_Progress.Label2.Caption = "STEP4/5 �W�v�\Excel�t�@�C���i�ڎ��j�쐬��..."
                Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
                waitTime = Now + TimeValue("0:00:01")
                Application.Wait waitTime
                
                ' �P���W�v���̕\�����̒���
                If face_flag = 0 Then
                    wb_spread.Activate
                    Sheets(Array("�m���\", "�m�\", "���\")).Select
                    ws_spread1.Columns(5).ColumnWidth = 12.37
                    ws_spread2.Columns(5).ColumnWidth = 12.37
                    ws_spread3.Columns(5).ColumnWidth = 12.37
                    Range("A1").Select
                End If
                
                ' �ڎ��̏���
                wb_spread.Activate
                ws_spread0.Select
                ActiveWindow.DisplayGridlines = False
                Cells.Select
                With Selection.Font
                    .Name = Mid(mcs_ini(2), 8)      ' ���{��t�H���g���Z�b�g
                    .Size = 9                       ' �t�H���g�T�C�Y���Z�b�g
                End With
                
                Columns(4).Select
                Columns(4).WrapText = True
                Columns(4).ColumnWidth = 42
                
                Columns(5).Select
                Columns(5).WrapText = True
                Columns(5).ColumnWidth = 42
                
                If WorksheetFunction.CountA(ws_spread0.Columns(6)) > 1 Then
                    Columns(6).Select
                    Columns(6).WrapText = True
                    Columns(6).ColumnWidth = 42
                End If
                ws_spread0.Columns(3).Hidden = True     ' MCODE��̔�\��
                
                ' �ڎ��̌r������
                Range("A1").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.LineStyle = xlContinuous
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                
                With Range(Selection, ActiveCell.SpecialCells(xlLastCell))
                    .Borders(xlInsideHorizontal).Weight = xlHairline
                End With
                With Range("A1:I1")
                    .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
                    .Borders(xlEdgeBottom).Weight = xlThin
                End With
                Range("G1:I1").Merge
                Range("G1").HorizontalAlignment = xlLeft
                Range("A2").Select
                ActiveWindow.FreezePanes = True
                Range("A1").Select
                
                ' �E�G�C�g�o�b�N�W�v���̐�����
                If ws_spread0.Cells(1, 1) = "�A�ԃE��" Then
                    ws_spread0.Cells(1, 1) = "�A��"
                    Range("A1").End(xlDown).Select
                    wgt_row = ActiveCell.Row
                    ws_spread0.Cells(wgt_row + 1, 1) = "���E�G�C�g�o�b�N�W�v���s���Ă��邽�߁A�v�Z�ߒ��ŏ����_�������܂����A�{�W�v�\��̐��l�͎l�̌ܓ����Đ����\�L���Ă��܂��B"
                    Range("A1").Select
                End If
                
                ' �e�V�[�g�̐F��ݒ�
                Sheets("�m���\").Select
                With ActiveWorkbook.Sheets("�m���\").Tab
                    .Color = 10066431
                    .TintAndShade = 0
                End With
                Sheets("�m�\").Select
                With ActiveWorkbook.Sheets("�m�\").Tab
                    .Color = 10092441
                    .TintAndShade = 0
                End With
                Sheets("���\").Select
                With ActiveWorkbook.Sheets("���\").Tab
                    .Color = 16764057
                    .TintAndShade = 0
                End With
                Sheets("�ڎ�").Select
                
                ' TOP1�EGT�\�[�g���� - 2018.10.04 �ǉ�
                ws_spread1.Select
                ws_spread1.PageSetup.Orientation = xlLandscape
                ActiveCell.SpecialCells(xlLastCell).Select
                max_row = ActiveCell.Row
                Range("A1").Select
                For i_cnt = 1 To hyo_cnt
                    DoEvents
                    Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
                    Form_Progress.Label2.Caption = "STEP5/5 �W�v�\Excel�t�@�C�� �ŏI������" & Status_Dot(i_cnt)
                    Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
                    Call top1_sort
                Next i_cnt
                Range("A1").Select
                wb_spread.Activate
                ws_spread0.Select
                
                Application.ScreenUpdating = True
                
                ' �W�v�\�T�}���[�t�@�C����ۑ����ăN���[�Y
                spread_fn = Replace(summary_fn, "sum", "�W�v�\")
                Open file_path & "\SUM\" & spread_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(spread_fn).Close
                End If
                If Dir(file_path & "\SUM\" & spread_fn) <> "" Then
                    Kill file_path & "\SUM\" & spread_fn
                End If
                
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=file_path & "\SUM\" & spread_fn
                ActiveWorkbook.Close
                Application.DisplayAlerts = True

                sum_cnt = sum_cnt + 1
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
            ' 2020.6.3 - �ǉ�
            ActiveSheet.Unprotect Password:=""
            ws_mainmenu.Cells(initial_row, initial_col).Locked = False
            If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
              ws_mainmenu.Cells(41, 6) = "25"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 25"
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
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�\Excel�t�@�C���̍쐬�F�Ώۃt�@�C���mSUM�t�H���_����" & sum_cnt - 1 & "�̏W�v�T�}���[�f�[�^�n"
            Close #1
            Call Finishing_Mcs2017
            MsgBox sum_cnt - 1 & "�̏W�v�\Excel�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Spreadsheet_Creation"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If
    
' �W�v�\Excel�t�@�C���P��쐬����
step00:
    wb.Activate
    ws_mainmenu.Select
    summary_fn = Application.GetOpenFilename("�W�v�T�}���[�t�@�C��,*.xlsx", , "�W�v�T�}���[�t�@�C�����J��")
    If summary_fn = "False" Then
        ' �L�����Z���{�^���̏���
        Call Finishing_Mcs2017
        End
    ElseIf summary_fn = "" Then
        MsgBox "�W�v�\Excel�t�@�C�����쐬����m�W�v�T�}���[�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        GoTo step00
    ElseIf InStr(summary_fn, "_sum") = 0 Then
        MsgBox "�W�v�\Excel�t�@�C�����쐬����m�W�v�T�}���[�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Spreadsheet_Creation"
        GoTo step00
    End If
    
    Open summary_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(summary_fn).Close
    Else
        Workbooks.Open summary_fn
    End If
    
    ' �t���p�X����t�H���_���̎擾
    yen_pos = InStrRev(summary_fn, "\")
    summary_fd = Left(summary_fn, yen_pos - 1)
    
    ' �t���p�X����t�@�C�����̎擾
    summary_fn = Dir(summary_fn)
    
    Set wb_spread = Workbooks(summary_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

    wb_spread.Activate
    Sheets(Array("�m���\", "�m�\", "���\")).Select
    Cells.Select
    
    ' �W�v�\�V�[�g�S�̂̃X�^�C���ݒ�
    ActiveWindow.DisplayGridlines = False
    With Selection.Font
        .Name = Mid(mcs_ini(2), 8)      ' ���{��t�H���g���Z�b�g
        .Size = Mid(mcs_ini(3), 13)     ' ���{��t�H���g�T�C�Y���Z�b�g
    End With
    
    ' �W�v�\�V�[�g�S�̂̕\���̃X�^�C���ݒ�
    Columns(1).Select
    With Selection
        .ColumnWidth = 4.88
    End With
    
    ' �W�v�\�V�[�g�S�̂̕\���J�e�S���[�ԍ��̃X�^�C���ݒ�
    Columns(3).Select
    With Selection
        .HorizontalAlignment = xlRight
        .ColumnWidth = 4.25
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With
    
    Range("A1").Select
    ws_spread0.Select
    ws_spread0.PageSetup.Orientation = xlLandscape
    hyo_cnt = ws_spread0.Cells(Rows.Count, setup_col).End(xlUp).Row - 1
    
    face_flag = 0       ' �N���X�W�v�\�̗L������p�t���O
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - �W�v�\Excel�t�@�C���̍쐬"
    Form_Progress.Repaint
    progress_msg = "�W�v�\Excel�t�@�C���̍쐬���L�����Z�����܂����B"
    Application.Visible = False
    AppActivate Form_Progress.Caption
    
    ' �m���\�̏���
    ws_spread1.Select
    ws_spread1.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP1/5 �W�v�\Excel�t�@�C���i�m���\�j�쐬��" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        Call int_spreadsheet
    Next i_cnt
    
    ws_spread1.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "���S�V�b�N"
      .Size = 8
    End With
    ws_spread1.Columns(2).Hidden = True     ' MCODE��̔�\��
    Range("A1").Select
    
    ' �m�\�̏���
    ws_spread2.Select
    ws_spread2.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
'---------------------------------
    np_max_row = ActiveCell.Row
'---------------------------------
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP2/5 �W�v�\Excel�t�@�C���i�m�\�j�쐬��" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        Call ken_spreadsheet
    Next i_cnt
    
    ws_spread2.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "���S�V�b�N"
      .Size = 8
    End With
    ws_spread2.Columns(2).Hidden = True     ' MCODE��̔�\��
    Range("A1").Select

    ' ���\�̏���
    ws_spread3.Select
    ws_spread3.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
'---------------------------------
    np_max_row = ActiveCell.Row
'---------------------------------
    
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP3/5 �W�v�\Excel�t�@�C���i���\�j�쐬��" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        Call per_spreadsheet
    Next i_cnt
    
    ws_spread3.Columns("B:B").Select
    Selection.ClearFormats
    With Selection.Font
      .Name = "���S�V�b�N"
      .Size = 8
    End With
    ws_spread3.Columns(2).Hidden = True     ' MCODE��̔�\��
    Range("A1").Select
    
    Form_Progress.Label1.Caption = "100%"
    DoEvents
    Form_Progress.Label2.Caption = "STEP4/5 �W�v�\Excel�t�@�C���i�ڎ��j�쐬��..."
    Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
    waitTime = Now + TimeValue("0:00:01")
    Application.Wait waitTime
    
    ' �P���W�v���̕\�����̒���
    If face_flag = 0 Then
        wb_spread.Activate
        Sheets(Array("�m���\", "�m�\", "���\")).Select
        ws_spread1.Columns(5).ColumnWidth = 12.37
        ws_spread2.Columns(5).ColumnWidth = 12.37
        ws_spread3.Columns(5).ColumnWidth = 12.37
        Range("A1").Select
    End If
    
    ' �ڎ��̏���
    wb_spread.Activate
    ws_spread0.Select
    ActiveWindow.DisplayGridlines = False
    Cells.Select
    With Selection.Font
        .Name = Mid(mcs_ini(2), 8)      ' ���{��t�H���g���Z�b�g
        .Size = 8                       ' �t�H���g�T�C�Y���Z�b�g
    End With
    
    Columns(4).Select
    Columns(4).WrapText = True
    Columns(4).ColumnWidth = 42
    
    Columns(5).Select
    Columns(5).WrapText = True
    Columns(5).ColumnWidth = 42
    
    If WorksheetFunction.CountA(ws_spread0.Columns(6)) > 1 Then
        Columns(6).Select
        Columns(6).WrapText = True
        Columns(6).ColumnWidth = 42
    End If
    ws_spread0.Columns(3).Hidden = True     ' MCODE��̔�\��
    
    ' �ڎ��̌r������
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.LineStyle = xlContinuous
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Borders.Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    
    With Range(Selection, ActiveCell.SpecialCells(xlLastCell))
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
    With Range("A1:I1")
        .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    Range("G1:I1").Merge
    Range("G1").HorizontalAlignment = xlLeft
    Range("A2").Select
    ActiveWindow.FreezePanes = True
    Range("A1").Select
    
    ' �E�G�C�g�o�b�N�W�v���̐�����
    If ws_spread0.Cells(1, 1) = "�A�ԃE��" Then
        ws_spread0.Cells(1, 1) = "�A��"
        Range("A1").End(xlDown).Select
        wgt_row = ActiveCell.Row
        ws_spread0.Cells(wgt_row + 1, 1) = "���E�G�C�g�o�b�N�W�v���s���Ă��邽�߁A�v�Z�ߒ��ŏ����_�������܂����A�{�W�v�\��̐��l�͎l�̌ܓ����Đ����\�L���Ă��܂��B"
        Range("A1").Select
    End If
    
    ' �e�V�[�g�̐F��ݒ�
    Sheets("�m���\").Select
    With ActiveWorkbook.Sheets("�m���\").Tab
        .Color = 10066431
        .TintAndShade = 0
    End With
    Sheets("�m�\").Select
    With ActiveWorkbook.Sheets("�m�\").Tab
        .Color = 10092441
        .TintAndShade = 0
    End With
    Sheets("���\").Select
    With ActiveWorkbook.Sheets("���\").Tab
        .Color = 16764057
        .TintAndShade = 0
    End With
    Sheets("�ڎ�").Select
    
' TOP1�EGT�\�[�g���� - 2018.10.04 �ǉ�
    ws_spread1.Select
    ws_spread1.PageSetup.Orientation = xlLandscape
    ActiveCell.SpecialCells(xlLastCell).Select
    max_row = ActiveCell.Row
    Range("A1").Select
    For i_cnt = 1 To hyo_cnt
        DoEvents
        Form_Progress.Label1.Caption = Int(i_cnt / hyo_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP5/5 �W�v�\Excel�t�@�C�� �ŏI������" & Status_Dot(i_cnt)
        Form_Progress.Label3.Caption = "[1/1�t�@�C��]"
        Call top1_sort
    Next i_cnt
    Range("A1").Select
    wb_spread.Activate
    ws_spread0.Select
    
    Application.ScreenUpdating = True
    
    ' �W�v�T�}���[�t�@�C����ۑ����ăN���[�Y
    spread_fn = Replace(summary_fn, "sum", "�W�v�\")
    Open file_path & "\SUM\" & spread_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(spread_fn).Close
    End If
    If Dir(file_path & "\SUM\" & spread_fn) <> "" Then
        Kill file_path & "\SUM\" & spread_fn
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=summary_fd & "\" & spread_fn
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Application.Visible = True
    Unload Form_Progress
    
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
      ws_mainmenu.Cells(41, 6) = "25"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 25"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�\Excel�t�@�C���̍쐬�F�Ώۃt�@�C���m" & summary_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�W�v�\Excel�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Spreadsheet_Creation"
End Sub

Private Sub int_spreadsheet()
' �����{���\�̃X�^�C���ݒ�
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread1.Select
    
    ' �W�v�\�P�\������̊J�n�s�擾
    bgn_row = ActiveCell.Row

    ' �W�v�\�P�\������̊J�n��擾
    bgn_col = ActiveCell.Column
    
    ' �W�v�\�P�\������̍ŏI�s�擾
    Selection.End(xlDown).Select
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' �������̍s��擾
    For k_cnt = bgn_row To fin_row
        If ws_spread1.Cells(k_cnt, 6) = "����" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread1.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column
    
    Selection.End(xlToRight).Select

    ' �W�v�\�P�\������̍ŏI��擾
    fin_col = ActiveCell.Column

    ' �\��̃X�^�C���ݒ�
    ws_spread1.Select
    With ws_spread1.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '�m�ݖ�`���n�Ɓm�\����ꐔ�n�̃X�^�C���ݒ� - 2018.05.24 �ǉ�
    Range(ws_spread1.Cells(ken_row, 4), ws_spread1.Cells(ken_row, 5)).Merge
    With ws_spread1.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' �\���J�e�S���[�ԍ��̃X�^�C���ݒ�
    With ws_spread1.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' �\�����ڂ̃X�^�C���ݒ�
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' �S�̍s�̃X�^�C���ݒ�
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "�@�S�@��" Then
        total_flag = 1
        ws_spread1.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread1.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' �S�̍s�̍s��擾
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread1.Cells(zen_row, zen_col), ws_spread1.Cells(zen_row + 1, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread1.Cells(zen_row, zen_col + 1), ws_spread1.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' �W�v�l�̃X�^�C���ݒ�
    With Range(ws_spread1.Cells(ken_row + 1, ken_col), ws_spread1.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' �p�����t�H���g���Z�b�g
        .Font.Size = Mid(mcs_ini(5), 13)     ' �p�����t�H���g�T�C�Y���Z�b�g
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' �\���\��̃X�^�C���ݒ�
    If ws_spread1.Cells(ken_row + 3, ken_col - 3) <> "" Then
        face_flag = 1
        cross_flag = 1
        
        ' �S�̗��̗L���ɂ���āA���W�𒲐�
        If total_flag = 1 Then
            adjust_row = 3
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread1.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread1.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread1.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread1.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' �W�v�\�P�\������̌r������
    With Range(ws_spread1.Cells(ken_row, ken_col - 2), ws_spread1.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread1.Cells(ken_row, ken_col - 2), ws_spread1.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread1.Cells(ken_row, ken_col), ws_spread1.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' �\�����ڂ̌r������
    If cross_flag = 1 Then
        For f_cnt = ken_row + 3 To fin_row - 1 Step 2
            With Range(ws_spread1.Cells(f_cnt, ken_col - 1), ws_spread1.Cells(f_cnt + 1, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' ���̏W�v�\�̕\���Ɉړ�
    ws_spread1.Cells(fin_row + 2, 1).Select

End Sub

Private Sub ken_spreadsheet()
' �����\�̃X�^�C���ݒ�
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer

    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread2.Select
    
    ' �W�v�\�P�\������̊J�n�s�擾
    bgn_row = ActiveCell.Row

    ' �W�v�\�P�\������̊J�n��擾
    bgn_col = ActiveCell.Column

    Selection.End(xlDown).Select

    ' �W�v�\�P�\������̍ŏI�s�擾
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' �������̍s��擾
    For k_cnt = bgn_row To fin_row
        If ws_spread2.Cells(k_cnt, 6) = "����" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread2.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    Selection.End(xlToRight).Select

    ' �W�v�\�P�\������̍ŏI��擾
    fin_col = ActiveCell.Column

    ' �\��̃X�^�C���ݒ�
    With ws_spread2.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '�m�ݖ�`���n�Ɓm�\����ꐔ�n�̃X�^�C���ݒ� - 2018.05.24 �ǉ�
    Range(ws_spread2.Cells(ken_row, 4), ws_spread2.Cells(ken_row, 5)).Merge
    With ws_spread2.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' �\���J�e�S���[�ԍ��̃X�^�C���ݒ�
    With ws_spread2.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' �\�����ڂ̃X�^�C���ݒ�
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' �S�̍s�̃X�^�C���ݒ�
    If ws_spread2.Cells(ken_row + 1, ken_col - 2) = "�@�S�@��" Then
        total_flag = 1
        ws_spread2.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread2.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' �S�̍s�̍s��擾
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread2.Cells(zen_row, zen_col), ws_spread2.Cells(zen_row, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread2.Cells(zen_row, zen_col + 1), ws_spread2.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' �W�v�l�̃X�^�C���ݒ�
    With Range(ws_spread2.Cells(ken_row + 1, ken_col), ws_spread2.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' �p�����t�H���g���Z�b�g
        .Font.Size = Mid(mcs_ini(5), 13)     ' �p�����t�H���g�T�C�Y���Z�b�g
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' �\���\��̃X�^�C���ݒ�
    If ws_spread2.Cells(ken_row + 2, ken_col - 3) <> "" Then
        cross_flag = 1
        
        ' �S�̗��̗L���ɂ���āA���W�𒲐�
        If total_flag = 1 Then
            adjust_row = 2
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread2.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread2.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread2.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread2.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' �W�v�\�P�\������̌r������
    With Range(ws_spread2.Cells(ken_row, ken_col - 2), ws_spread2.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread2.Cells(ken_row, ken_col - 2), ws_spread2.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread2.Cells(ken_row, ken_col), ws_spread2.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' �\�����ڂ̌r������
    If cross_flag = 1 Then
        For f_cnt = ken_row + 2 To fin_row Step 1
            With Range(ws_spread2.Cells(f_cnt, ken_col - 1), ws_spread2.Cells(f_cnt, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' ���̏W�v�\�̕\���Ɉړ�
    ws_spread2.Cells(fin_row + 2, 1).Select

End Sub

Private Sub per_spreadsheet()
' ���\�̃X�^�C���ݒ�
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    Dim adjust_row As Integer
    Dim adjust_col As Integer
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread3.Select
    
    ' �W�v�\�P�\������̊J�n�s�擾
    bgn_row = ActiveCell.Row

    ' �W�v�\�P�\������̊J�n��擾
    bgn_col = ActiveCell.Column

    Selection.End(xlDown).Select

    ' �W�v�\�P�\������̍ŏI�s�擾
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' �������̍s��擾
    For k_cnt = bgn_row To fin_row
        If ws_spread3.Cells(k_cnt, 6) = "����" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread3.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    Selection.End(xlToRight).Select

    ' �W�v�\�P�\������̍ŏI��擾
    fin_col = ActiveCell.Column

    ' �\��̃X�^�C���ݒ�
    With ws_spread3.Cells(bgn_row, 4)
        .Font.Size = 11
        .Font.Bold = True
    End With

    '�m�ݖ�`���n�Ɓm�\����ꐔ�n�̃X�^�C���ݒ� - 2018.05.24 �ǉ�
    Range(ws_spread3.Cells(ken_row, 4), ws_spread3.Cells(ken_row, 5)).Merge
    With ws_spread3.Cells(ken_row, 4)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
    End With

    ' �\���J�e�S���[�ԍ��̃X�^�C���ݒ�
    With ws_spread3.Rows(ken_row - 1)
        .HorizontalAlignment = xlLeft
        .Font.Color = RGB(0, 112, 192)
        .Font.Italic = True
        .Font.Size = 7
    End With

    ' �\�����ڂ̃X�^�C���ݒ�
    With Range(Cells(ken_row, ken_col), Cells(ken_row, fin_col))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .RowHeight = 90
        .ColumnWidth = 10.5
    End With

    ' �S�̍s�̃X�^�C���ݒ�
    If ws_spread3.Cells(ken_row + 1, ken_col - 2) = "�@�S�@��" Then
        total_flag = 1
        ws_spread3.Cells(ken_row + 1, ken_col - 2).Select
        ws_spread3.Cells(ken_row + 1, ken_col - 2).ColumnWidth = 10.63
    
        ' �S�̍s�̍s��擾
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        With Range(ws_spread3.Cells(zen_row, zen_col), ws_spread3.Cells(zen_row, fin_col))
            .Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
        With Range(ws_spread3.Cells(zen_row, zen_col + 1), ws_spread3.Cells(zen_row, zen_col + 1))
            .ColumnWidth = 48.13
        End With
    End If

    ' �W�v�l�̃X�^�C���ݒ�
    With Range(ws_spread3.Cells(ken_row + 1, ken_col), ws_spread3.Cells(fin_row, fin_col))
        .Font.Name = Mid(mcs_ini(4), 8)      ' �p�����t�H���g���Z�b�g
        .Font.Size = Mid(mcs_ini(5), 13)     ' �p�����t�H���g�T�C�Y���Z�b�g
        .HorizontalAlignment = xlRight
        .ShrinkToFit = True
    End With

    ' �\���\��̃X�^�C���ݒ�
    If ws_spread3.Cells(ken_row + 2, ken_col - 3) <> "" Then
        cross_flag = 1
        
        ' �S�̗��̗L���ɂ���āA���W�𒲐�
        If total_flag = 1 Then
            adjust_row = 2
            adjust_col = -2
        Else
            adjust_row = 1
            adjust_col = -2
        End If
        
        Range(ws_spread3.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread3.Cells(fin_row, ken_col + adjust_col)).Merge
        With Range(ws_spread3.Cells(ken_row + adjust_row, ken_col + adjust_col), ws_spread3.Cells(fin_row, ken_col + adjust_col))
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        End With
    End If

    ' �W�v�\�P�\������̌r������
    With Range(ws_spread3.Cells(ken_row, ken_col - 2), ws_spread3.Cells(fin_row, fin_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread3.Cells(ken_row, ken_col - 2), ws_spread3.Cells(ken_row, fin_col))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With
    With Range(ws_spread3.Cells(ken_row, ken_col), ws_spread3.Cells(fin_row, fin_col))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
    End With

    ' �\�����ڂ̌r������
    If cross_flag = 1 Then
        For f_cnt = ken_row + 2 To fin_row Step 1
            With Range(ws_spread3.Cells(f_cnt, ken_col - 1), ws_spread3.Cells(f_cnt, fin_col))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            End With
        Next f_cnt
    End If

    ' ���̏W�v�\�̕\���Ɉړ�
    ws_spread3.Cells(fin_row + 2, 1).Select

End Sub

Private Sub top1_sort()
' TOP1�EGT�\�[�g���� - 2018.11.13 �ŏI�X�V
    Dim bgn_row As Long
    Dim bgn_col As Long
    Dim fin_row As Long
    Dim fin_col As Long
    Dim ken_row As Long
    Dim ken_col As Long
    Dim zen_row As Long
    Dim zen_col As Long
'---------------------------------
    Dim val_range As Range
    Dim c_cnt As Long
    Dim cx_cnt As Long
    Dim face_cnt As Long
    Dim ct_row As Long
    Dim ct_col As Long
    Dim ct_cnt As Long
    Dim top1_ct As Double
    Dim top1_col As Long
    Dim ex_ct As Long
    Dim h_num As String
    Dim np_bgn_row As Long
    Dim np_bgn_col As Long
    Dim np_fin_row As Long
    Dim np_fin_col As Long
    Dim np_ken_row As Long
    Dim np_ken_col As Long
    Dim np_zen_row As Long
    Dim np_zen_col As Long
    Dim np_ct_row As Long
    Dim np_ct_col As Long
'---------------------------------
    Dim total_flag As Integer
    Dim cross_flag As Integer
    Dim f_cnt As Long
    Dim k_cnt As Long
    
    total_flag = 0
    cross_flag = 0
    
    wb_spread.Activate
    ws_spread1.Select
    
    ' �W�v�\�P�\������̊J�n�s�擾
    bgn_row = ActiveCell.Row

    ' �W�v�\�P�\������̊J�n��擾
    bgn_col = ActiveCell.Column
    
    ' �W�v�\�P�\������̍ŏI�s�擾
    Selection.End(xlDown).Select
    If ActiveCell.Row = 1048576 Then
        fin_row = max_row
    Else
        fin_row = ActiveCell.Row - 2
    End If
    Selection.End(xlUp).Select

    ' �������̍s��擾
    For k_cnt = bgn_row To fin_row
        If ws_spread1.Cells(k_cnt, 6) = "����" Then
            ken_row = k_cnt
            Exit For
        End If
    Next k_cnt
    ws_spread1.Cells(ken_row, 6).Select
    ken_col = ActiveCell.Column

    ' �J�e�S���[�ԍ��̍s��擾
    ct_row = ken_row - 1
    ct_col = ken_col + 1
    
    Selection.End(xlToRight).Select

    ' �W�v�\�P�\������̍ŏI��擾
    fin_col = ActiveCell.Column

    ' �W�v�\�P�\������̏����ݒ�
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "�@�S�@��" Then
        ' �S�̍s�̍s��擾
        ws_spread1.Cells(ken_row + 1, ken_col - 2).Select
        zen_row = ActiveCell.Row
        zen_col = ActiveCell.Column
        
        ' �m�\�E���\�̍��W�擾
        h_num = ws_spread1.Cells(bgn_row, bgn_col)
        ws_spread2.Select
        Columns("A:A").Select
        Selection.Find(What:=h_num, after:=ActiveCell, LookIn:=xlFormulas, _
         lookat:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
         MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Select
        np_bgn_row = ActiveCell.Row
        np_bgn_col = ActiveCell.Column
    
        ' �m�\�E���\�̏W�v�\�P�\������̍ŏI�s�擾
        Selection.End(xlDown).Select
        If ActiveCell.Row = 1048576 Then
            np_fin_row = np_max_row
        Else
            np_fin_row = ActiveCell.Row - 2
        End If
        Selection.End(xlUp).Select
    
        ' �m�\�E���\�̌������̍s��擾
        For k_cnt = np_bgn_row To np_fin_row
            If ws_spread2.Cells(k_cnt, 6) = "����" Then
                np_ken_row = k_cnt
                Exit For
            End If
        Next k_cnt
        ws_spread2.Cells(np_ken_row, 6).Select
        np_ken_col = ActiveCell.Column

        ' �m�\�E���\�̃J�e�S���[�ԍ��̍s��擾
        np_ct_row = np_ken_row - 1
        np_ct_col = np_ken_col + 1
        Range("A1").Select
    End If

    ' GT�\�[�g�̏��� - 2018.10.10
    If Mid(ws_spread1.Cells(bgn_row + 2, bgn_col + 1), 1, 1) = "Y" Then
        ws_spread1.Select
        ct_cnt = 0
        For c_cnt = ct_col To fin_col    ' �J�e�S���[�����J�E���g
            If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                Exit For
            ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                Exit For
            End If
            ct_cnt = ct_cnt + 1
        Next c_cnt

        ' ���OCT�̊m�F
        ex_ct = Val(Mid(ws_spread1.Cells(bgn_row + 2, bgn_col + 1), 2))

        ' ���OCT��CT���𒴂���ꍇ�̓\�[�g�������Ȃ�
        If ct_cnt >= ex_ct Then
            If ex_ct <> 0 Then
                ct_cnt = ex_ct - 1
            End If

            ' �\�[�g����
            Range(ws_spread1.Cells(ken_row - 1, ken_col + 1), ws_spread1.Cells(fin_row, ken_col + ct_cnt)).Select
            ws_spread1.Sort.SortFields.Clear
            ws_spread1.Sort.SortFields.Add Key:=Rows(zen_row), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            With ws_spread1.Sort
                .SetRange Range(ws_spread1.Cells(ken_row - 1, ken_col + 1), ws_spread1.Cells(fin_row, ken_col + ct_cnt))
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlLeftToRight
                .SortMethod = xlStroke
                .Apply
            End With
            
            ' �m���\����m�\�E���\�֐U�蕪��
            For c_cnt = 0 To ct_cnt
                ws_spread2.Cells(np_ken_row - 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row - 1, ken_col + c_cnt)
                ws_spread3.Cells(np_ken_row - 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row - 1, ken_col + c_cnt)
                ws_spread2.Cells(np_ken_row, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row, ken_col + c_cnt)
                ws_spread3.Cells(np_ken_row, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row, ken_col + c_cnt)
                If c_cnt = 0 Then
                    ws_spread2.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                    ws_spread3.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                Else
                    ws_spread2.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                    ws_spread3.Cells(np_ken_row + 1, np_ken_col + c_cnt) = ws_spread1.Cells(ken_row + 1, ken_col + c_cnt)
                End If
            Next c_cnt
        
            ' �N���X�W�v�\�̏c�W�J����
            If face_flag = 1 Then
                ' �m���\����m�\�֐U�蕪��
                face_cnt = 1
                For cx_cnt = zen_row + 2 To fin_row - 1 Step 2
                    For c_cnt = 0 To ct_cnt
                        If c_cnt = 0 Then
                            ws_spread2.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        Else
                            ws_spread2.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        End If
                    Next c_cnt
                    face_cnt = face_cnt + 1
                Next cx_cnt
                ' �m���\���灓�\�֐U�蕪��
                face_cnt = 1
                For cx_cnt = zen_row + 2 To fin_row - 1 Step 2
                    For c_cnt = 0 To ct_cnt
                        If c_cnt = 0 Then
                            ws_spread3.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        Else
                            ws_spread3.Cells(np_ken_row + 1 + face_cnt, np_ken_col + c_cnt) = ws_spread1.Cells(cx_cnt, ken_col + c_cnt)
                        End If
                    Next c_cnt
                    face_cnt = face_cnt + 1
                Next cx_cnt
            End If
        End If
    End If

    ' �S�̍s�̃X�^�C���ݒ�
    If ws_spread1.Cells(ken_row + 1, ken_col - 2) = "�@�S�@��" Then
        ' TOP1�i�S�́j�̏��� - 2018.09.28 �ǉ�
        ws_spread1.Select
        If ws_spread1.Cells(bgn_row + 1, bgn_col + 1) <> "" Then
            ct_cnt = 0
            For c_cnt = ct_col To fin_col    ' �J�e�S���[�����J�E���g
                If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                    Exit For
                ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                    Exit For
                End If
                ct_cnt = ct_cnt + 1
            Next c_cnt
            ws_spread1.Select
            If ct_cnt <> 0 Then    ' ct_cnt���m0�n�Ȃ�A�ΏۂȂ��i�J�e�S���[���Ȃ��ݖ�c���肦�Ȃ��Ǝv�����ǁc�j
                Set val_range = Range(ws_spread1.Cells(ct_row + 2, ct_col), ws_spread1.Cells(ct_row + 2, ct_col + ct_cnt - 1))
                top1_ct = Application.WorksheetFunction.Max(val_range)
                Set val_range = Nothing
                For top1_col = ct_col To (ct_col + ct_cnt - 1)
                    If ws_spread1.Cells(ct_row + 2, top1_col) = top1_ct Then
                        If ws_spread1.Cells(ct_row + 2, top1_col).Value <> 0 Then    '������ �m0���n�Ȃ璅�F���Ȃ�
                            Range(ws_spread1.Cells(ct_row + 2, top1_col), ws_spread1.Cells(ct_row + 3, top1_col)).Interior.Color = 8420607
                            Range(ws_spread2.Cells(np_ct_row + 2, top1_col), ws_spread2.Cells(np_ct_row + 2, top1_col)).Interior.Color = 8420607
                            Range(ws_spread3.Cells(np_ct_row + 2, top1_col), ws_spread3.Cells(np_ct_row + 2, top1_col)).Interior.Color = 8420607
                        End If
                    End If
                Next top1_col
            End If
        End If
    End If
    
    ' �\���\��̃X�^�C���ݒ�
    If ws_spread1.Cells(ken_row + 3, ken_col - 3) <> "" Then
        ' TOP1�i�\�����ځj�̏��� - 2018.10.03 �ǉ�
        ct_cnt = 0
        For c_cnt = ct_col To fin_col    ' �J�e�S���[�����J�E���g
            If ws_spread1.Cells(ct_row, c_cnt) = "N/A" Then
                Exit For
            ElseIf ws_spread1.Cells(ct_row, c_cnt) = "" Then
                Exit For
            End If
            ct_cnt = ct_cnt + 1
        Next c_cnt
        If ct_cnt <> 0 Then    ' ct_cnt���m0�n�Ȃ�A�ΏۂȂ��i�J�e�S���[���Ȃ��ݖ�c���肦�Ȃ��Ǝv�����ǁc�j
            If ws_spread1.Cells(bgn_row + 1, bgn_col + 1) = "A" Then
                If face_cnt = 0 Then
                    face_cnt = 1
                    For f_cnt = zen_row + 3 To fin_row Step 2
                        face_cnt = face_cnt + 1
                    Next f_cnt
                End If
                For f_cnt = 1 To face_cnt - 1
                    If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), zen_col - 1) <> "N/A" Then
                        Set val_range = Range(ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), ct_col), ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), ct_col + ct_cnt - 1))
                        top1_ct = Application.WorksheetFunction.Max(val_range)
                        Set val_range = Nothing
                        For top1_col = ct_col To (ct_col + ct_cnt - 1)
                            If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col) = top1_ct Then
                                If ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col).Value <> 0 Then    '������ �m0���n�Ȃ璅�F���Ȃ�
                                    Range(ws_spread1.Cells(ct_row + 2 + (f_cnt * 2), top1_col), ws_spread1.Cells(ct_row + 3 + (f_cnt * 2), top1_col)).Interior.Color = 8420607
                                    Range(ws_spread2.Cells(np_ct_row + 2 + f_cnt, top1_col), ws_spread2.Cells(np_ct_row + 2 + f_cnt, top1_col)).Interior.Color = 8420607
                                    Range(ws_spread3.Cells(np_ct_row + 2 + f_cnt, top1_col), ws_spread3.Cells(np_ct_row + 2 + f_cnt, top1_col)).Interior.Color = 8420607
                                End If
                            End If
                        Next top1_col
                    End If
                Next f_cnt
            End If
        End If
    End If

    ' ���̏W�v�\�̕\���Ɉړ�
    ws_spread1.Cells(fin_row + 2, 1).Select

End Sub

