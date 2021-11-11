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

    ' �C���w���̃A�h���X�ݒ�
    Const e_sno As Integer = 1      ' SampleNo
    Const e_qcode As Integer = 2    ' QCODE
    Const e_data As Integer = 5     ' �񓚓��e
    Const e_rst  As Integer = 6     ' �C�����e
'--------------------------------------------------------------------------------------------------'
'�@���̓f�[�^�̏C���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�c���@�`�W�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.05.15�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
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
    rev_fn = Application.GetOpenFilename("�C���w���t�@�C��,*.xlsx", , "�C���w���t�@�C�����J��")
    If rev_fn = "False" Then
        ' �L�����Z���{�^���̏���
        Call Finishing_Mcs2017
        End
    ElseIf rev_fn = "" Then
        MsgBox "�m�C���w���t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Indata_Creation"
        GoTo step00
    ElseIf InStr(rev_fn, "_�C���w��.xlsx") = 0 Then
        MsgBox "�m�C���w���t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Indata_Creation"
        GoTo step00
    End If
    
    Open rev_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(Dir(rev_fn)).Close
    Else
        Workbooks.Open rev_fn
    End If
    
    ' �t���p�X����t�@�C�����̎擾
    rev_fn = Dir(rev_fn)
    
    Set wb_revision = Workbooks(rev_fn)
    Set ws_revision = wb_revision.Worksheets(1)
    
    wb_revision.Activate
    ws_revision.Select
    rev_row = Cells(Rows.Count, 1).End(xlUp).Row
    Application.StatusBar = False
        
    ws_revision.Cells(5, 7).Value = "�C���O"
    ws_revision.Cells(5, 8).Value = "�C����"

    idata_fn = ws_revision.Cells(2, 4)
    odata_fn = ws_revision.Cells(3, 4)
    
    ChDrive file_path & "\1_DATA"
    ChDir file_path & "\1_DATA"
    
    ' �C���t�@�C���̗L���`�F�b�N
    If Dir(file_path & "\1_DATA\" & idata_fn) = "" Then
        MsgBox "�C���t�@�C�����Őݒ肳��Ă���t�@�C���m" & idata_fn & "�n��������܂���B", vbExclamation, "MCS 2020 - Indata_Creation"
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
    
    ' �o�̓t�@�C���̗L���`�F�b�N
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
        
        ' �f�[�^�̃t�H�[�}�b�g�m�F
        If Cells(1, 1) <> "SNO" Then
            MsgBox "�T���v���i���o�[��QCODE�ɁmSNO�n�ȊO���ݒ肳��Ă��܂��B" & vbCrLf & "�C���w���t�@�C���́m�C���t�@�C�����n���m�F���Ă��������B", vbExclamation, "MCS 2020 - Indata_Revision"
            End
        End If
        
        ' �C���ΏۃT���v���i���o�[�������A��������s���擾
        Set FoundCell = Range(Cells(7, 1), ws_data.Cells(max_row, 1)).Find(What:=rev_sno, lookat:=xlWhole)
        If FoundCell Is Nothing Then
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = "�w�肵���T���v���i���o�[���f�[�^��Ō�����܂���B"
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "�C���ł��܂���ł����B"
        Else
            dat_row = FoundCell.Row
        End If

        ' �C���Ώۍ��ځiQCODE�j�������A�����������擾
        Set FoundCell = Range(Cells(1, 1), ws_data.Cells(1, max_col)).Find(What:=rev_qcode, lookat:=xlWhole)
        If FoundCell Is Nothing Then
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = "�Ώۂ�QCODE��������܂���B"
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "�C���ł��܂���ł����B"
        Else
            dat_col = FoundCell.Column
            If rev_mact <> 0 Then
                dat_col = dat_col + rev_mact - 1
            End If
        End If

        ' �C�����e�irev_after�j���A�w�u�����N�iNULL�j�x�A�w�N���A�x�A�w���񓚁x�Ȃ�
        If (rev_after = "") Or (rev_after = "�N���A") Or (rev_after = "����") Then
            ' �C���O�̃f�[�^���o��
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = ws_data.Cells(dat_row, dat_col).Value
            ' �f�[�^���C��
            ws_data.Cells(dat_row, dat_col).Value = ""
            ' �C����̃f�[�^���o��
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "�N���A�i���񓚁j"
            
        ' �C�����e�irev_after�j�����l�Ȃ�
        ElseIf IsNumeric(rev_after) = True Then
            ' �C���O�̃f�[�^���o��
            ws_revision.Cells(rev_cnt, e_rst + 1).Value = ws_data.Cells(dat_row, dat_col).Value
            ' �f�[�^���C��
            ws_data.Cells(dat_row, dat_col).Value = rev_after
            ' �C����̃f�[�^���o��
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = ws_data.Cells(dat_row, dat_col).Value
            
        ' �C�����e�irev_after�j��"DEL"�Ȃ�
        ElseIf rev_after = "DEL" Then
            ' �f�[�^���C��
            ws_data.Cells(dat_row, dat_col).EntireRow.Delete
            ' �C����̃f�[�^���o��
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "�T���v���J�b�g"
            
        ' �C�����e�irev_after�j����L�ȊO�Ȃ�
        Else
            ws_revision.Cells(rev_cnt, e_rst + 2).Value = "������"
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
    
' �V�X�e�����O�̏o��
    ' 2020.6.3 - �ǉ�
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - ���̓f�[�^�̏C���F�g�p�t�@�C���m" & rev_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�f�[�^�̏C�����������܂����B", vbInformation, "MCS 2020 - Indata_Revision"
End Sub

