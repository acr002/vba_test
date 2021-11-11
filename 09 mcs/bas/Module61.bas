Attribute VB_Name = "Module61"
Option Explicit
    Dim mcs_ini(10) As String
    Dim ini_cnt As Integer

    Dim spread_fn As String
    Dim spread_fd As String
    
    Dim rep_fn As String
    Dim rep_fd As String
    
    Dim wb_spread As Workbook
    Dim ws_spread As Worksheet
    
    Dim wb_report As Workbook
    Dim ws_report As Worksheet
    Dim ws_value As Worksheet

'2020/04/17 - �ǋL ==========================
    Dim graph_ptn() As Integer

Sub Simplicity_report()
    Dim rc As Integer
    Dim yen_pos As Long
    Dim max_row As Long, max_col As Long

'2018/06/01 - �ǋL ==========================
    Dim r_code As Integer
    Dim spd_tab() As String
    Dim spd_file As String
    Dim spd_cnt As Long
    Dim n_cnt As Long
    Dim fn_cnt As Long

'2020/01/07 - �ǋL ==========================
    Dim i_cnt As Long, r_cnt As Long
    Dim a_cnt As Long
    Dim h_row() As Long
    Dim v_row As Long, v_col As Long
    Dim r_row As Long, r_col As Long
    Dim sel_cnt As Long, ans_cnt As Long
    
'2020/04/10 - �ǋL ==========================
    Dim num_format As String
'--------------------------------------------------------------------------------------------------'
'�@�P���W�v���|�[�g�t�@�C���̍쐬�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.07.09�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.20�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "�P���W�v���|�[�g�t�@�C���̍쐬��..."
    
    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    wb.Activate
    ws_mainmenu.Select
    MsgBox "�P���W�v�̏W�v�T�}���[�f�[�^��I�����Ă��������B"
step00:
    spread_fn = Application.GetOpenFilename("�W�v�T�}���[�f�[�^,*.xlsx", , "�P���W�v�̏W�v�T�}���[�f�[�^���J��")
    If spread_fn = "False" Then
        ' �L�����Z���{�^���̏���
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf spread_fn = "" Then
        MsgBox "�P���W�v�̏W�v�T�}���[�f�[�^��I�����Ă��������B", vbExclamation, "MCS 2020 - Simplicity_report"
        GoTo step00
    ElseIf InStr(spread_fn, "_sum.xlsx") = 0 Then
        MsgBox "�P���W�v�̏W�v�T�}���[�f�[�^��I�����Ă��������B", vbExclamation, "MCS 2020 - Simplicity_report"
        GoTo step00
    End If

    Open spread_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(Dir(spread_fn)).Close
    Else
        Workbooks.Open spread_fn
        If Application.WorksheetFunction.Sum(Columns(3)) <> 0 Then
            Workbooks(Dir(spread_fn)).Close
            MsgBox "�P���W�v�̏W�v�T�}���[�f�[�^��I�����Ă��������B", vbExclamation, "MCS 2020 - Simplicity_report"
            wb.Activate
            ws_mainmenu.Select
            GoTo step00
        End If
    End If
    
    ' �t���p�X����t�H���_���̎擾
    yen_pos = InStrRev(spread_fn, "\")
    spread_fd = Left(spread_fn, yen_pos - 1)
    
    ' �t���p�X����t�@�C�����̎擾
    spread_fn = Dir(spread_fn)
    
    ' �t�@�C��������g���q�ȊO�̎擾
    rep_fn = Left(spread_fn, InStr(spread_fn, "_sum") - 1)
    
    Set wb_spread = Workbooks(spread_fn)
    Set ws_spread = wb_spread.Worksheets("�m���\")

' �P���W�v���|�[�g�t�@�C���쐬��������
    Application.DisplayAlerts = False
    wb_spread.Activate
    
    rep_fd = file_path & "\SUM\Report"
    If Dir(rep_fd, vbDirectory) = "" Then
        MkDir rep_fd
    End If
    
    Workbooks.Add
    Set wb_report = ActiveWorkbook
    
    ActiveSheet.Name = "���|�[�g"
    Worksheets.Add after:=ActiveSheet
    ActiveSheet.Name = "�W�v�l"
    
    Set ws_report = wb_report.Worksheets("���|�[�g")
    Set ws_value = wb_report.Worksheets("�W�v�l")
    
    ' �W�v�T�}���[�f�[�^�����|�[�g�t�@�C���̏W�v�l�V�[�g�ɃR�s�y
    wb_spread.Activate
    Cells.Select
    Selection.Copy
    wb_report.Activate
    ActiveSheet.Paste
    Range("A1").Select
    wb_spread.Close
    
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - �P���W�v���|�[�g�t�@�C���̍쐬"
    Form_Progress.Repaint
    progress_msg = "�P���W�v���|�[�g�t�@�C���̍쐬���L�����Z�����܂����B"
    Application.Visible = False
    AppActivate Form_Progress.Caption

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
    
    ' ���|�[�g�V�[�g�̏����ݒ�
    ws_report.Select
    Cells.Select
    With Selection.Font
        .Name = Mid(mcs_ini(2), 8)      ' ���{��t�H���g���Z�b�g
        .Size = Mid(mcs_ini(3), 13)     ' ���{��t�H���g�T�C�Y���Z�b�g
    End With
    Range("A1").Select
    Range("B1").ColumnWidth = 30
    Range("C:D").ColumnWidth = 7
    Range("E1").ColumnWidth = 1
    Columns(2).WrapText = True
    
    ' �W�v�\�̐��ƕ\�ԍ��̍s���擾
    ws_value.Select
    r_cnt = Application.WorksheetFunction.CountA(ws_value.Range("A:A"))
    ReDim h_row(r_cnt)
    v_row = 1
    For i_cnt = 1 To r_cnt
      ws_value.Cells(v_row, 1).Select
      If ws_value.Cells(v_row, 1) <> "" Then
        h_row(i_cnt) = ActiveCell.Row
        Selection.End(xlDown).Select
        v_row = ActiveCell.Row
      End If
    Next i_cnt
    ws_value.Range("A1").Select
    
    ReDim graph_ptn(r_cnt)
    
    ' ���|�[�g�V�[�g�ɐ�����ݒ�
    r_row = 1: r_col = 1
    For i_cnt = 1 To r_cnt
      DoEvents
      
      Form_Progress.Label1.Caption = Int(i_cnt / r_cnt * 100) & "%"
      Form_Progress.Label2.Caption = "STEP1/2 �P���W�v���|�[�g�t�@�C���쐬��" & Status_Dot(i_cnt)
'      Form_Progress.Label3.Caption = "[" & sum_cnt & "/" & fn_cnt & "�t�@�C��]"
      
      v_row = h_row(i_cnt)
      v_col = 4
      ' �\��̐ݒ�
      If ws_value.Cells(v_row, 1) <> "" Then    ' �\�ԍ��̗L�����`�F�b�N
        ' �\��̍s�ԍ����擾
        ws_report.Cells(r_row, r_col).Formula = "=�W�v�l!" & ColNum2Let(v_col) & v_row
        ws_report.Cells(r_row, r_col).Font.Size = 11
        ws_report.Cells(r_row, r_col).Font.Bold = True
        v_row = v_row + 1
        r_row = r_row + 1
    
        For sel_cnt = 1 To 7
        ' �Z���N�g�̐ݒ�@�`�F
          If ws_value.Cells(v_row, v_col) <> "" Then
            ws_report.Cells(r_row, r_col).Formula = "=�W�v�l!" & ColNum2Let(v_col) & v_row
            v_row = v_row + 1
            r_row = r_row + 1
          Else
            Exit For
          End If
        Next sel_cnt
    
        ' �����̎��_�ł� v_row �́A�W�v�l�V�[�g�̏W�v�\�̕\���̃u�����N�Z����I�����Ă���B
      
        ' �O���t�p�^�[���̎擾
        graph_ptn(i_cnt) = ws_value.Cells(v_row + 2, v_col - 2)
        
        ' �ݖ�`���E�\����ꐔ�̏���
        ws_report.Cells(r_row, r_col).Formula = "=�W�v�l!" & ColNum2Let(v_col) & v_row + 1
        r_row = r_row + 1
      
        ' �Œ茩�o���̏���
        ws_report.Cells(r_row, r_col + 2) = "����"
        ws_report.Cells(r_row, r_col + 2).HorizontalAlignment = xlCenter
        ws_report.Cells(r_row, r_col + 3) = "�\����"
        ws_report.Cells(r_row, r_col + 3).HorizontalAlignment = xlCenter
        ws_report.Rows(r_row).RowHeight = 18
        r_row = r_row + 1
    
        ' �S�̂̏���
        ws_report.Cells(r_row, r_col + 1) = "�S��"
        ws_report.Cells(r_row, r_col + 2).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2) & v_row + 2
        ws_report.Cells(r_row, r_col + 3) = "100"
        ws_report.Cells(r_row, r_col + 3).NumberFormatLocal = "0.0"
        ws_report.Rows(r_row).RowHeight = 25.5
        r_row = r_row + 1
     
        ' �\�����ڂ̏���
        ans_cnt = 0
        For a_cnt = 7 To 306    ' �J�e�S���[���i�S�̏����A���񓚂Ɠ��v�ʊ܂ށj�̎擾
          If ws_value.Cells(v_row + 1, a_cnt) <> "" Then
            ans_cnt = ans_cnt + 1
          Else
            Exit For
          End If
        Next a_cnt
    
        For a_cnt = 1 To ans_cnt    ' �\�����ڂ̓W�J�i���񓚊܂ށj
          If ws_value.Cells(v_row, a_cnt + 6) <> "" Then
            ' �����ƍ\����̏���
            ws_report.Rows(r_row).RowHeight = 25.5
            ws_report.Cells(r_row, r_col).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row
            ws_report.Cells(r_row, r_col).HorizontalAlignment = xlRight
            ws_report.Cells(r_row, r_col + 1).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row + 1
            ws_report.Cells(r_row, r_col + 1).HorizontalAlignment = xlLeft
            ws_report.Cells(r_row, r_col + 2).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row + 2
            ws_report.Cells(r_row, r_col + 2).HorizontalAlignment = xlRight
            ws_report.Cells(r_row, r_col + 2).NumberFormatLocal = "0"
            ws_report.Cells(r_row, r_col + 3).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row + 3
            ws_report.Cells(r_row, r_col + 3).HorizontalAlignment = xlRight
            ws_report.Cells(r_row, r_col + 3).NumberFormatLocal = "0.0"
          Else
            ' ���v�ʂ̏���
            ws_report.Rows(r_row).RowHeight = 18
            ws_report.Cells(r_row, r_col + 1).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row + 1
            ws_report.Cells(r_row, r_col + 1).HorizontalAlignment = xlLeft
            ws_report.Cells(r_row, r_col + 2).Formula = "=�W�v�l!" & ColNum2Let(v_col + 2 + a_cnt) & v_row + 2
            Range(ws_report.Cells(r_row, r_col + 2), ws_report.Cells(r_row, r_col + 3)).MergeCells = True
            num_format = ws_value.Cells(v_row + 2, v_col + 2 + a_cnt).NumberFormatLocal
            ws_report.Cells(r_row, r_col + 2).NumberFormatLocal = num_format
            ws_report.Cells(r_row, r_col + 2).HorizontalAlignment = xlRight
          End If
          r_row = r_row + 1
        Next a_cnt
      
        r_row = r_row + 1
      End If
    Next i_cnt
    
    Form_Progress.Label2.Caption = "STEP2/2 �P���W�v���|�[�g�t�@�C���ŏI������..."
    Call Report_Keisen
    Call Report_Graph(r_cnt)
    Call Page_Setup

    ActiveWorkbook.SaveAs Filename:=rep_fd & "\" & rep_fn & "_Report.xlsx"
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
' �P���W�v���|�[�g�t�@�C���쐬�����܂�
    
    Application.Visible = True
    Unload Form_Progress
    Application.ScreenUpdating = True
    
    Set wb_spread = Nothing
    Set ws_spread = Nothing
    Set wb_report = Nothing
    Set ws_report = Nothing
    Set ws_value = Nothing

    wb.Activate
    ws_setup.Select
    ws_setup.Cells(1, 1).Select
    ws_mainmenu.Select
    ws_mainmenu.Cells(3, 8).Select
    
' �V�X�e�����O�̏o��
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "24"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 24"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �P���W�v���|�[�g�t�@�C���̍쐬�F�Ώۃt�@�C���m" & spread_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�P���W�v���|�[�g�t�@�C�����������܂����B", vbInformation, "MCS 2020 - Simplicity_report"
End Sub

Private Sub Report_Keisen()
' �r������
    Dim last_row As Long
    Dim i_cnt As Long
    Dim k_row As Long
    
    ws_report.Select
    last_row = ws_report.Cells(Rows.Count, 2).End(xlUp).Row

    For i_cnt = 2 To last_row
        DoEvents
        If (ws_report.Cells(i_cnt, 2) = "�S��") And (ws_report.Cells(i_cnt - 1, 3) = "����") Then
            ws_report.Cells(i_cnt, 2).Select
            Selection.End(xlDown).Select
            k_row = ActiveCell.Row
            
            Range(ws_report.Cells(i_cnt - 1, 2), ws_report.Cells(k_row, 4)).Select
            Selection.Borders.Color = RGB(Mid(mcs_ini(7), 14, 3), Mid(mcs_ini(7), 18, 3), Mid(mcs_ini(7), 22, 3))
            Selection.Borders.LineStyle = True
            
            Range(ws_report.Cells(i_cnt, 2), ws_report.Cells(i_cnt, 4)).Select
            Selection.Interior.Color = RGB(Mid(mcs_ini(6), 13, 3), Mid(mcs_ini(6), 17, 3), Mid(mcs_ini(6), 21, 3))
        
            ws_report.Cells(i_cnt - 1, 2).Select
            Selection.Borders(xlEdgeTop).LineStyle = False
            Selection.Borders(xlEdgeLeft).LineStyle = False
        End If
    Next i_cnt
    
    Range("A1").Select
End Sub

Private Sub Report_Graph(ByVal rx_cnt As Long)
' �O���t����
    Dim last_row As Long
    Dim i_cnt As Long, g_cnt As Long
    Dim bgn_row As Long, btm_row As Long
    Dim ct_info As Long, na_info As Long
    Dim graph_cls() As Long
    
    ws_report.Select
    last_row = ws_report.Cells(Rows.Count, 1).End(xlUp).Row

    g_cnt = 1
    ReDim graph_cls(rx_cnt)
    For i_cnt = 2 To last_row
        DoEvents
        If (ws_report.Cells(i_cnt, 2) = "�S��") And (ws_report.Cells(i_cnt - 1, 3) = "����") Then
            ws_report.Cells(i_cnt + 1, 1).Select
            bgn_row = ActiveCell.Row
            Selection.End(xlDown).Select
            
            ' �O���t�͈͂Ɂm���񓚁n���܂ނ��̔���
            If InStr(ws_report.Cells(i_cnt - 2, 1).Value, "�\����ꐔ�F�S��") > 0 Then
                If ws_report.Cells(ActiveCell.Row - 1, 1) = "" Then
                    ' �P�J�e�S���[�̂Ƃ��̏���
                    btm_row = bgn_row
                    ct_info = 1
                    na_info = 0
                Else
                    btm_row = ActiveCell.Row
                    If ws_report.Cells(ActiveCell.Row, 1) = "N/A" Then
                        ct_info = ws_report.Cells(ActiveCell.Row - 1, 1)
                        na_info = 1
                    Else
                        ct_info = ws_report.Cells(ActiveCell.Row, 1)
                        na_info = 0
                    End If
                End If
            Else
                If ws_report.Cells(ActiveCell.Row, 1) = "N/A" Then
                    btm_row = ActiveCell.Row - 1
                    ct_info = ws_report.Cells(ActiveCell.Row - 1, 1)
                    na_info = 1
                Else
                    If ws_report.Cells(ActiveCell.Row - 1, 1) = "" Then
                        ' �P�J�e�S���[�̂Ƃ��̏���
                        btm_row = bgn_row
                        ct_info = 1
                        na_info = 0
                    Else
                        btm_row = ActiveCell.Row
                        ct_info = ws_report.Cells(ActiveCell.Row, 1)
                        na_info = 0
                    End If
                End If
            End If
        
            graph_cls(g_cnt) = 0
            Select Case graph_ptn(g_cnt)
            
            ' �~�O���t�̏���
            Case 1
                If ct_info = 1 Then
                    ' �P�J�e�S���[�̏ꍇ�A�`��̈悪���������߉��_�O���t���쐬
                    Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
                ElseIf (ct_info + na_info) > 6 Then
                    ' �U�J�e�S���[���̏ꍇ�A�J�e�S���[�����������߉��_�O���t���쐬
                    Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
                Else
                    Call Graph_Pie(g_cnt, bgn_row, btm_row, ct_info, na_info)
                End If
            
            ' ���_�O���t�̏���
            Case 2
                Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
            
            ' ���Ė_�O���t�̏���
            Case 3
                Call Graph_ColumnClustered(g_cnt, bgn_row, btm_row)
            
            ' �уO���t�̏���
            Case 4
                If ct_info = 1 Then
                    ' �P�J�e�S���[�̏ꍇ�A�`��̈悪���������߉��_�O���t���쐬
                    Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
                ElseIf (ct_info + na_info) > 6 Then
                    ' �U�J�e�S���[���̏ꍇ�A�J�e�S���[�����������߉��_�O���t���쐬
                    Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
                Else
                    Call Graph_BarStacked100(g_cnt, bgn_row, btm_row, ct_info, na_info)
                End If
            
            ' �O���t�p�^�[����v�Ȃ�
            Case Else
                ' �_�~�[�ŉ��_�O���t�쐬��A�폜�t���O���I���ɂ���B
                Call Graph_BarClustered(g_cnt, bgn_row, btm_row)
                graph_cls(g_cnt) = 1
            End Select
        
            g_cnt = g_cnt + 1
        End If
    Next i_cnt
    
    ' �O���t�p�^�[���s��v�ɂ��O���t�폜����
    For i_cnt = 1 To g_cnt - 1
        If graph_cls(i_cnt) = 1 Then
            ActiveSheet.ChartObjects(i_cnt).Activate
            ActiveChart.Parent.Delete
        End If
    Next i_cnt
    
    Range("A1").Select
End Sub

Private Sub Graph_Pie(ByVal gx_cnt As Long, ByVal bgnx_row As Long, ByVal btmx_row As Long, _
 ByVal ctx_info As Long, ByVal nax_info As Long)
' �~�O���t�̏���
    Dim l_cnt As Long
    Dim all_info As Long
    Dim height_adj As Double
    Dim rr As Integer, gg As Integer, bb As Integer
    Dim rng As Range

    With ActiveSheet.ChartObjects.Add(400, 40, 350, 200).Chart
        .ChartType = xlPie
        .SetSourceData Source:=Range(ws_report.Cells(bgnx_row, 4), ws_report.Cells(btmx_row, 4))
        .ChartArea.Border.LineStyle = 0
        .HasLegend = True
    End With
    
    all_info = ctx_info + nax_info
    Select Case all_info
    Case 2
        ' �Q�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 3, 5), ws_report.Cells(btmx_row + 2, 10))
        height_adj = 15.02362205
    Case 3
        ' �Q�J�e�S���[�{���񓚁A�R�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 3, 5), ws_report.Cells(btmx_row + 1, 10))
        height_adj = 7.653543312
    Case 4
        ' �R�J�e�S���[�{���񓚁A�S�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 2, 5), ws_report.Cells(btmx_row, 10))
        height_adj = 7.37007874
    Case 5
        ' �S�J�e�S���[�{���񓚁A�T�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row, 10))
        height_adj = 0
    Case 6
        ' �T�J�e�S���[�{���񓚁i�����U�J�e�S���[���j
        Set rng = Range(ws_report.Cells(bgnx_row, 5), ws_report.Cells(btmx_row, 10))
        height_adj = 0
    Case Else
        ' ���Y���f�[�^�Ȃ�
    End Select
    
    With ActiveSheet.ChartObjects(gx_cnt)
        .Top = rng.Top - height_adj
        .Left = rng.Left - 5
        .Width = rng.Width
        .Height = rng.Height + height_adj
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
    
    ActiveChart.SeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(204, 236, 255)
        .Transparency = 0
    End With
    
    ' �J�e�S���[�����Ƃ̔z�F
    Select Case all_info
    Case 2
        ' �Q�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
                        
        ActiveChart.SeriesCollection(1).Points(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Transparency = 0
            .Solid
        End With
    Case 3
        ' �Q�J�e�S���[�{���񓚁A�R�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
                        
        ActiveChart.SeriesCollection(1).Points(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 51: gg = 102: bb = 255
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(1).Points(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
    Case 4
        ' �R�J�e�S���[�{���񓚁A�S�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 102, 255)
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 102: gg = 153: bb = 255
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(1).Points(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
    Case 5
        ' �S�J�e�S���[�{���񓚁A�T�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 102, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(102, 153, 255)
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 153: gg = 204: bb = 255
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(1).Points(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
    Case 6
        ' �T�J�e�S���[�{���񓚁i�����U�J�e�S���[���j
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 102, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(102, 153, 255)
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(1).Points(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(153, 204, 255)
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 204: gg = 204: bb = 255
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(1).Points(6).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
    Case Else
        ' �������Y���f�[�^�͂Ȃ��͂��B
        ActiveChart.SeriesCollection(1).Points(1).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(51, 51, 255)
            .Solid
        End With
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 0, 102)
            .Transparency = 0
            .Solid
        End With
    End Select
    
    ' �f�[�^���x���̏���
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SeriesCollection(1).DataLabels.Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    With Selection.Format.TextFrame2.TextRange.Font
        .BaselineOffset = 0
        .NameFarEast = "���S�V�b�N"
        .Size = 8
        .Name = "���S�V�b�N"
    End With
    Selection.Position = xlLabelPositionInsideEnd
    
    For l_cnt = 1 To all_info
        ActiveChart.SeriesCollection(1).Points(l_cnt).DataLabel.Select
        Selection.Position = xlLabelPositionBestFit
    Next l_cnt
    
    ' �}��̏���
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SeriesCollection(1).XValues = Range(ws_report.Cells(bgnx_row, 1), ws_report.Cells(btmx_row, 1))
    
    ActiveChart.Legend.Select
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "���S�V�b�N"
        .NameFarEast = "���S�V�b�N"
        .Name = "���S�V�b�N"
    End With
    Selection.Format.TextFrame2.TextRange.Font.Size = 8

    Set rng = Nothing
End Sub

Private Sub Graph_BarClustered(ByVal gx_cnt As Long, ByVal bgnx_row As Long, ByVal btmx_row As Long)
' ���_�O���t�̏���
    Dim rng As Range
    
    With ActiveSheet.ChartObjects.Add(400, 40, 350, 200).Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=Range(ws_report.Cells(bgnx_row, 4), ws_report.Cells(btmx_row, 4))
        .ChartArea.Border.LineStyle = 0
        .HasLegend = False
        .Axes(xlCategory).TickLabelPosition = xlNone
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).MajorTickMark = xlTickMarkInside
        .Axes(xlValue).MajorTickMark = xlTickMarkNone
        .Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineSysDot
        .Axes(xlValue).MaximumScale = 100
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).TickLabels.NumberFormat = """�@""0""% """
        .Axes(xlValue).TickLabels.Font.Name = "Arial"
        .Axes(xlValue).TickLabels.Font.Size = 7
    End With
    
    Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row, 10))
    With ActiveSheet.ChartObjects(gx_cnt)
        .Top = rng.Top + 9
        .Left = rng.Left - 5
        .Width = rng.Width
        .Height = rng.Height - 8
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
    
    ActiveChart.PlotArea.Select
    ActiveChart.PlotArea.Border.LineStyle = xlContinuous
    Selection.Top = Selection.Top - 9
    Selection.Left = Selection.Left - 20
    Selection.Height = Selection.Height + 19.5
    Selection.Width = Selection.Width + 30
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(137, 137, 137)
        .Transparency = 0
    End With
    
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.ChartGroups(1).GapWidth = 50
    Selection.Format.Line.Visible = msoFalse
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .ForeColor.RGB = RGB(0, 102, 255)    ' �_�O���t�̐F�ݒ�
        .Transparency = 0
        .Solid
    End With

    Set rng = Nothing
End Sub

Private Sub Graph_ColumnClustered(ByVal gx_cnt As Long, ByVal bgnx_row As Long, ByVal btmx_row As Long)
' ���Ė_�O���t�̏���
    Dim rng As Range
    Dim plot_rng As Range
    
    With ActiveSheet.ChartObjects.Add(400, 40, 350, 200).Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=Range(ws_report.Cells(bgnx_row, 4), ws_report.Cells(btmx_row, 4))
        .ChartArea.Border.LineStyle = 0
        .HasLegend = False
        .Axes(xlCategory).ReversePlotOrder = False
        .Axes(xlCategory).Crosses = xlMaximum
        .Axes(xlCategory).TickLabels.Font.Name = "���S�V�b�N"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).MajorTickMark = xlTickMarkNone
        .Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineSysDot
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).MaximumScaleIsAuto = True
        .Axes(xlValue).TickLabels.NumberFormat = "0""% """
        .Axes(xlValue).TickLabels.Font.Name = "Arial"
        .Axes(xlValue).TickLabels.Font.Size = 7
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    Selection.Delete
    ActiveChart.Axes(xlValue).MajorTickMark = xlInside
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
    ActiveChart.Axes(xlCategory).Select
    Selection.MajorTickMark = xlNone
    ActiveChart.SeriesCollection(1).XValues = Range(ws_report.Cells(bgnx_row, 1), ws_report.Cells(btmx_row, 1))
    
    Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row, 10))
    With ActiveSheet.ChartObjects(gx_cnt)
        .Top = rng.Top
        .Left = rng.Left
        .Width = rng.Width
        .Height = rng.Height
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
    
    Set plot_rng = Range(ws_report.Cells(bgnx_row, 6), ws_report.Cells(btmx_row, 10))
    ActiveChart.PlotArea.Select
    ActiveChart.PlotArea.Border.LineStyle = xlContinuous
    Selection.Top = ws_report.Cells(bgnx_row, 6).Top
    Selection.Left = Selection.Left - 1
    Selection.Height = plot_rng.Height - 5
    Selection.Width = Selection.Width
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(137, 137, 137)
        .Transparency = 0
    End With
    
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.ChartGroups(1).GapWidth = 50
    Selection.Format.Line.Visible = msoFalse
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .ForeColor.RGB = RGB(0, 102, 255)    ' �_�O���t�̐F�ݒ�
        .Transparency = 0
        .Solid
    End With

    ' �f�[�^���x���̏���
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SeriesCollection(1).DataLabels.Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 255, 255)
        .Solid
    End With
    
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "���S�V�b�N"
        .NameFarEast = "���S�V�b�N"
        .Name = "���S�V�b�N"
    End With
    Selection.Format.TextFrame2.TextRange.Font.Size = 8

    Set rng = Nothing
    Set plot_rng = Nothing
End Sub

Private Sub Graph_BarStacked100(ByVal gx_cnt As Long, ByVal bgnx_row As Long, ByVal btmx_row As Long, _
 ByVal ctx_info As Long, ByVal nax_info As Long)
' �уO���t�̏���
    Dim v_flag() As String
    Dim l_cnt As Long, v_cnt As Long
    Dim a_val As Double, b_val As Double, c_val As Double
    Dim height_adj As Double
    Dim plot_adj As Double
    Dim all_info As Long
    Dim rr As Integer, gg As Integer, bb As Integer
    Dim rng As Range

    With ActiveSheet.ChartObjects.Add(400, 40, 350, 200).Chart
        .ChartType = xlBarStacked100
        .SetSourceData Source:=Range(ws_report.Cells(bgnx_row, 4), ws_report.Cells(btmx_row, 4))
        .ChartArea.Border.LineStyle = 0
        .HasLegend = True
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    With ActiveChart
        Select Case .PlotBy
        Case xlRows
            .PlotBy = xlColumns
        Case xlColumns
            .PlotBy = xlRows
        End Select
    End With
    
    ActiveChart.Axes(xlCategory).Select
    Selection.MajorTickMark = xlNone
    ActiveChart.Axes(xlCategory).Crosses = xlMaximum
    Selection.Delete

    all_info = ctx_info + nax_info
    Select Case all_info
    Case 2
        ' �Q�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 3, 5), ws_report.Cells(btmx_row + 1, 10))
        height_adj = 7.653543307
        plot_adj = 7
    Case 3
        ' �Q�J�e�S���[�{���񓚁A�R�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row + 1, 10))
        height_adj = 12.75590551
        plot_adj = 12.5
    Case 4
        ' �R�J�e�S���[�{���񓚁A�S�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row, 10))
        height_adj = 0
        plot_adj = 0
    Case 5
        ' �S�J�e�S���[�{���񓚁A�T�J�e�S���[�i���񓚗��\���Ȃ��j
        Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row - 1, 10))
        height_adj = 0
        plot_adj = 0
    Case 6
        ' �T�J�e�S���[�{���񓚁i�����U�J�e�S���[���j
        Set rng = Range(ws_report.Cells(bgnx_row - 1, 5), ws_report.Cells(btmx_row - 2, 10))
        height_adj = 0
        plot_adj = 0
    Case Else
        ' ���Y���f�[�^�Ȃ�
    End Select
    
    With ActiveSheet.ChartObjects(gx_cnt)
        .Top = rng.Top - height_adj
        .Left = rng.Left
        .Left = rng.Left - 10
        .Width = rng.Width + 5
        .Height = rng.Height + height_adj
    End With
    
    ActiveSheet.ChartObjects(gx_cnt).Activate
    ActiveChart.ChartArea.Interior.ColorIndex = xlNone
    ActiveChart.Axes(xlValue).Select
    Selection.MajorTickMark = xlNone
    Selection.MinimumScale = 0
    Selection.MaximumScale = 1
    Selection.TickLabels.NumberFormatLocal = """�@""0%"
    Selection.TickLabels.Font.Name = "Arial"
    Selection.TickLabels.Font.Size = 7
    
    ' �v���b�g�G���A�̔�����
    ActiveChart.PlotArea.Select
    ActiveChart.PlotArea.Border.LineStyle = xlContinuous
    Selection.Top = Selection.Top
    Selection.Left = Selection.Left - 5
    Selection.Height = rng.Height - 36 + plot_adj
    Selection.Width = rng.Width
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(137, 137, 137)
        .Transparency = 0
    End With

    ActiveChart.Axes(xlValue).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = -0.150000006
        .Transparency = 0
    End With
    
    ' �}��̐ݒ�
    ActiveChart.Legend.Position = xlBottom
    ActiveChart.Legend.Select
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "���S�V�b�N"
        .NameFarEast = "���S�V�b�N"
        .Name = "���S�V�b�N"
    End With
    Selection.Format.TextFrame2.TextRange.Font.Size = 8
    
    For l_cnt = 1 To all_info
        ActiveChart.SeriesCollection(l_cnt).Name = ws_report.Cells(bgnx_row + l_cnt - 1, 1)
    Next l_cnt
    
    ' �J�e�S���[�����Ƃ̔z�F
    Select Case all_info
    Case 2
        ' �Q�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 204, 0)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(2).ApplyDataLabels
        ActiveChart.SeriesCollection(2).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    Case 3
        ' �Q�J�e�S���[�{���񓚁A�R�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 204, 0)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(2).ApplyDataLabels
        ActiveChart.SeriesCollection(2).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 153: gg = 255: bb = 102
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(3).ApplyDataLabels
        ActiveChart.SeriesCollection(3).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    Case 4
        ' �R�J�e�S���[�{���񓚁A�S�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 204, 0)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(2).ApplyDataLabels
        ActiveChart.SeriesCollection(2).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(153, 255, 102)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(3).ApplyDataLabels
        ActiveChart.SeriesCollection(3).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 102: gg = 255: bb = 153
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(4).ApplyDataLabels
        ActiveChart.SeriesCollection(4).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    Case 5
        ' �S�J�e�S���[�{���񓚁A�T�J�e�S���[�i���񓚗��\���Ȃ��j
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 204, 0)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(2).ApplyDataLabels
        ActiveChart.SeriesCollection(2).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(153, 255, 102)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(3).ApplyDataLabels
        ActiveChart.SeriesCollection(3).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(102, 255, 153)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(4).ApplyDataLabels
        ActiveChart.SeriesCollection(4).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 204: gg = 255: bb = 204
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(5).ApplyDataLabels
        ActiveChart.SeriesCollection(5).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    Case 6
        ' �T�J�e�S���[�{���񓚁i�����U�J�e�S���[���j
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(2).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 204, 0)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(2).ApplyDataLabels
        ActiveChart.SeriesCollection(2).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(3).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(153, 255, 102)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(3).ApplyDataLabels
        ActiveChart.SeriesCollection(3).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(4).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(102, 255, 153)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(4).ApplyDataLabels
        ActiveChart.SeriesCollection(4).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        ActiveChart.SeriesCollection(5).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(204, 255, 204)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(5).ApplyDataLabels
        ActiveChart.SeriesCollection(5).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
        
        If nax_info = 0 Then
            rr = 234: gg = 255: bb = 234
        Else
            rr = 255: gg = 255: bb = 255
        End If
        
        ActiveChart.SeriesCollection(6).Select
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(rr, gg, bb)
            .Transparency = 0
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(6).ApplyDataLabels
        ActiveChart.SeriesCollection(6).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    Case Else
        ' �������Y���f�[�^�͂Ȃ��͂��B
        ActiveChart.SeriesCollection(1).Select
        ActiveChart.ChartGroups(1).GapWidth = 100
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 153, 0)
            .Solid
        End With
        With Selection.Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(137, 137, 137)
            .Transparency = 0
        End With
        ActiveChart.SeriesCollection(1).ApplyDataLabels
        ActiveChart.SeriesCollection(1).DataLabels.Select
        With Selection.Format.TextFrame2.TextRange.Font
            .NameComplexScript = "���S�V�b�N"
            .NameFarEast = "���S�V�b�N"
            .Name = "���S�V�b�N"
        End With
        Selection.Format.TextFrame2.TextRange.Font.Size = 8
        With Selection.Format.TextFrame2.TextRange.Font.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0
            .Transparency = 0
            .Solid
        End With
    End Select
    
    ' ���l���x���̒��� - 2020.4.30���C
    v_cnt = 1
    ReDim v_flag(all_info)
    For l_cnt = 1 To all_info
        v_flag(v_cnt) = ""
        If ws_report.Cells(bgnx_row + (v_cnt - 1), 4).Value < 5 Then  ' ���x���l�T�������Ȃ珈���Ώ�
            ' �擪�J�e�S���[�̏���
            If l_cnt = 1 Then
                a_val = ws_report.Cells(bgnx_row + (v_cnt - 1), 4).Value    ' ����
                b_val = ws_report.Cells(bgnx_row + v_cnt, 4).Value          ' ��
                If (b_val >= 5) And ((a_val + b_val) < 11) Then
                    ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                    ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                    Selection.Top = Selection.Top - 27
                    v_flag(v_cnt) = "Up"
                End If
            ' �ŏI�J�e�S���[�̏���
            ElseIf l_cnt = all_info Then
                a_val = ws_report.Cells(bgnx_row + (v_cnt - 1), 4).Value    ' ����
                b_val = ws_report.Cells(bgnx_row + (v_cnt - 2), 4).Value    ' ��O
                If v_flag(v_cnt - 1) = "Right" Then
                    ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                    ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                    Selection.Top = Selection.Top + 27
                    v_flag(v_cnt) = "Down"
                ElseIf (v_flag(v_cnt - 1) = "") And (b_val < 1) Then    ' ��O���P���������ƃL�c�C�̂ŉ��ֈړ�
                    ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                    ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                    Selection.Top = Selection.Top + 27
                    v_flag(v_cnt) = "Down"
                ElseIf (a_val + b_val) < 11 Then    ' ��O�̃��x���l�Ƃ̍��v�l���P�P�������Ȃ�E�ֈړ�
                    ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                    ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                    Selection.Left = Selection.Left + 10
                    v_flag(v_cnt) = "Right"
                End If
            ' �r���J�e�S���[�̏���
            Else
                a_val = ws_report.Cells(bgnx_row + (v_cnt - 1), 4).Value    ' ����
                b_val = ws_report.Cells(bgnx_row + (v_cnt - 2), 4).Value    ' ��O
                c_val = ws_report.Cells(bgnx_row + v_cnt, 4).Value          ' ��
                
                ' ��O�̃��x���̏󋵂��m�F
                If v_flag(v_cnt - 1) = "" Then
                    If b_val < 5 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top + 27
                        v_flag(v_cnt) = "Down"
                    ElseIf c_val < 5 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top - 27
                        v_flag(v_cnt) = "Up"
                    ElseIf (a_val + c_val) < 11 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top - 27
                        v_flag(v_cnt) = "Up"
                    End If
                ElseIf v_flag(v_cnt - 1) = "Up" Then
                    If (c_val >= 5) And (a_val + c_val) < 11 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top + 27
                        v_flag(v_cnt) = "Down"
                    End If
                ElseIf v_flag(v_cnt - 1) = "Down" Then
                    If (a_val + b_val) < 11 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Left + 10
                        v_flag(v_cnt) = "Right"
                    Else
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top - 27
                        v_flag(v_cnt) = "Up"
                    End If
                Else
                    If (c_val >= 5) And (a_val + c_val) < 11 Then
                        ActiveChart.SeriesCollection(l_cnt).DataLabels.Select
                        ActiveChart.SeriesCollection(l_cnt).Points(1).DataLabel.Select
                        Selection.Top = Selection.Top - 27
                        v_flag(v_cnt) = "Top"
                    End If
                End If
            End If
        End If
        v_cnt = v_cnt + 1
    Next l_cnt

    Set rng = Nothing
End Sub

Private Sub Page_Setup()
' �y�[�W�ݒ�
    Dim last_row As Long
    Dim l_cnt As Long, b_cnt As Long
    Dim pb_cnt As Long
    Dim pb_row As Integer
    Dim prt_height As Double
    
    ws_report.Select
    last_row = ws_report.Cells(Rows.Count, 2).End(xlUp).Row
    
    Application.PrintCommunication = False
    ActiveSheet.PageSetup.PrintArea = "$A:$J"
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftMargin = Application.CentimetersToPoints(1#)
        .RightMargin = Application.CentimetersToPoints(1#)
        .TopMargin = Application.CentimetersToPoints(1#)
        .BottomMargin = Application.CentimetersToPoints(1#)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 95
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With

    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    
    prt_height = 0
    pb_cnt = 1
    For l_cnt = 1 To 1048576
        DoEvents
        If pb_cnt > last_row Then Exit For
        prt_height = prt_height + Range("A" & pb_cnt).Height
        If prt_height > 894 Then
            For b_cnt = pb_cnt To 1 Step -1
                pb_row = Application.WorksheetFunction.CountA(Rows(pb_cnt & ":" & pb_cnt))
                If pb_row = 0 Then
                    Rows(pb_cnt + 1).PageBreak = xlPageBreakManual
                    Exit For
                Else
                    pb_cnt = pb_cnt - 1
                End If
            Next b_cnt
            prt_height = 0
        End If
        pb_cnt = pb_cnt + 1
    Next l_cnt

    ActiveWindow.View = xlNormalView
    With ActiveSheet.PageSetup
        .CenterFooter = "&""���S�V�b�N,�W��""&9- &P -"
    End With
    Application.PrintCommunication = True
End Sub

Function ColNum2Let(ByVal colNum As Long, Optional colStr As String = "") As String
  ' ��ԍ����A���t�@�x�b�g�ɕϊ�
    If colNum = 0 Then
        ColNum2Let = colStr
    Else
        colStr = Chr(65 + (colNum - 1) Mod 26) & colStr
        colNum = (colNum - 1) \ 26
        ColNum2Let = ColNum2Let(colNum, colStr)
    End If
End Function
