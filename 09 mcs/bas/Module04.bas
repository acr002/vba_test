Attribute VB_Name = "Module04"
Option Explicit

' �e���H�G���A��ԍ��i�萔�j
Private Enum COLUMN_DATA

    SKIP_FLG = 3
    
    SF_START = 4
    
    QCODE1 = 5
    QCODE1_DATA1 = 6
    QCODE1_DATA2 = 7
    QCODE1_DATA3 = 8
    QCODE1_DATA4 = 9
    QCODE1_DATA5 = 10
    QCODE1_DATA6 = 11
    QCODE1_DATA7 = 12
    QCODE1_DATA8 = 13
    QCODE1_DATA9 = 14
    QCODE1_DATA10 = 15
    
    QCODE2 = 20
    QCODE2_DATA1 = 21
    QCODE2_DATA2 = 22
    QCODE2_DATA3 = 23
    QCODE2_DATA4 = 24
    QCODE2_DATA5 = 25

End Enum

' �e���H�G���A��ԍ��i�萔�j
Private Enum row_data

    START_ROW = 6
    
    ' �Z���N�g�t���O����p
    QCODE_P_CA = 6
    QCODE_P_ROW = 5
    QCODE_P_COLUMN = 4

End Enum

' �e���H�G���A��ԍ��i�萔�j
Private Enum ROW_INDATA

    START_ROW_INDATA = 7
    START_SUTATUSBER = 5
    
    START_TABLE_DATA = 6

End Enum

Public Sub processing_indata()
'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2017.05.18  '
' ���̓f�[�^���H�v���V�[�W�����C�����[�`��                                                         '
'--------------------------------------------------------------------------------------------------'
    Dim order_count As Long         ' ���H�������Ԕ���p�ϐ�
    Dim order_max As Long           ' ���H�����񐔎擾�p�ϐ�
    
    Dim wb_process As Workbook      ' ���H�w�����[�N�u�b�N�i�[�p�I�u�W�F�N�g
    Dim ws_menu As Worksheet        ' ���H�w�����C�����j���[�i�[�p�I�u�W�F�N�g
    Dim ws_process As Worksheet     ' ���H�w����ƃ��[�N�V�[�g�i�[�p�I�u�W�F�N�g
    
    Dim work_process As Long        ' ��Ɖ��H���i�[�p�ϐ�
    
    Dim wsp_indata As Worksheet      ' ���̓f�[�^���[�N�V�[�g�i�[�p�I�u�W�F�N�g
    Dim indata_maxrow As Long       ' ���̓f�[�^�ő�񐔊i�[�p�ϐ�
    
    Dim indata_maxcolumn As Long    ' ���̓f�[�^�ő�s���i�[�p�ϐ�
    Dim indata_count As Long        ' ���̓f�[�^�w�b�_���J�E���g�p�ϐ�
    Dim hedder_flg As Boolean       ' �w�b�_���t���O�i*���H��j
    Dim hedder_address As String    ' �w�b�_�ʒu�A�h���X�i�[�p�ϐ�
    
    Dim statusBar_text As String    ' �X�e�[�^�X�o�[�R�����g�i�[�p�ϐ�
    
    Dim connect_data As String      ' �����Z���N�g��������p�ϐ�

    Dim filename_work As String     ' ���H�w���t�@�C�����i�[�p�ϐ�
    
    Dim error_tb As Workbook        ' �G���[�o�͗p���[�N�u�b�N�I�u�W�F�N�g
    Dim error_ts As Worksheet       ' �G���[�o�͗p���[�N�V�[�g�I�u�W�F�N�g
    
    Dim indata_logname As String    ' ���O�o�͗p�C���f�[�^�l�[��
    
'    Dim now_data As Databar         ' �����i�[�p�ϐ�

    Dim check_wb As Workbook        ' �I�[�v���`�F�b�N�p���[�N�u�b�N�I�u�W�F�N�g
    Dim check_flg As Boolean        ' �I�[�v���`�F�b�N���ʊi�[�p�t���O
    Dim check_name As String        ' �I�[�v���`�F�b�N�p�t�@�C�����i�[�p�ϐ�
    
    Call Indata_Open
    Call Setup_Hold
    Call Filepath_Get
    Call Setup_Check
    
    indata_logname = wb_indata.Name
    
    ' ��ʂւ̕\�����I�t�ɂ���
    Application.ScreenUpdating = False
    
    ' ���b�Z�[�W���\���ɂ���
    Application.DisplayAlerts = False
    
    ' ���̓f�[�^�����[�N�V�[�g�Ƃ��Ċm��
    Set wsp_indata = wb_indata.Worksheets(1)
    
    ' ���̓f�[�^�ő�s
    indata_maxcolumn = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' �w�b�_�����i�[
    For indata_count = 1 To indata_maxcolumn
        ' �w�b�_����*���H�オ���邩�𔻒�
        If wsp_indata.Cells(1, indata_count) = "*���H��" Then
            hedder_flg = True
        End If
    Next
    
    ' �w�b�_��*���H�オ�܂܂�Ă��Ȃ������ꍇ
    If hedder_flg = False Then
        
        ' �擪�̃A�h���X���擾
        hedder_flg = Hedder_Create(wsp_indata, "*���H��", wsp_indata.Cells(1, indata_count).Address)
        
'        ' �R�����g����̓f�[�^�w�b�_�ɑ���A����������
'        wsp_indata.Range(hedder_address).Value = "*���H��"
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'
'        ' �w�b�_�A�h���X��ύX
'        hedder_address = wsp_indata.Range(hedder_address).Offset(1).Resize(3).Address
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
'
'        ' �S�̂Ɍr��������
'        hedder_address = wsp_indata.Range(hedder_address).Offset(3).Resize(2).Address
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlDash
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).Weight = xlHairline
'
'        ' ��ɐF������
'        wsp_indata.Columns(indata_count).Interior.Color = RGB(255, 255, 0)
        
    End If
    
    ' ���̓f�[�^�ő�񐔂��i�[
    indata_maxrow = wsp_indata.Cells(Rows.Count, 1).End(xlUp).Row
    
' �t�@�C�����擾����
step00:

    ChDrive file_path & "\3_FD"
    ChDir file_path & "\3_FD"
    
    filename_work = Application.GetOpenFilename("���H�w���t�@�C��,*.xlsm", , "���H�w���t�@�C�����J��")
    If filename_work = "False" Then
        ' �L�����Z���{�^���̏���
        End
    ElseIf filename_work = "" Then
        MsgBox "�m���H�w���t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2017 - processing_indata"
        GoTo step00
    ElseIf InStr(filename_work, "���H�w��") = 0 Then
        MsgBox "�m���H�w���t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2017 - processing_indata"
        GoTo step00
    End If
    
    ' �t�@�C�����݂̂𒊏o
    check_name = Mid(filename_work, Len(file_path & "\3_FD\"))
    
    ' �t�@�C�����J���Ă��邩���`�F�b�N
    For Each check_wb In Workbooks
        If check_wb.Name = check_name Then
            check_flg = True
        End If
    Next check_wb
    
    ' �u�b�N���J���Ă��邩����
    If check_flg = True Then
        Set ws_menu = wb_process.Worksheets("���C�����j���[")
        
    Else
        Set wb_process = Workbooks.Open(filename_work)
        Set ws_menu = wb_process.Worksheets("���C�����j���[")
    
    End If
    
    ' ���H�񐔂��擾
    order_max = ws_menu.Range("AE29").Value
    
    ' 20180614
    ' ���H���O���e�o�͗p�t�@�C�����m�F
    Set error_tb = Workbooks.Add
    Set error_ts = error_tb.Worksheets(1)
    
    ' ���H�o�͂̃w�b�_���쐬
    error_ts.Range("A1").Value = "SEQ"
    error_ts.Range("B1").Value = "���H���e"
    error_ts.Range("C1").Value = "QCode1"
    error_ts.Range("D1").Value = "QCode2"
    error_ts.Range("E1").Value = "�������e"
    
    ' ������
    error_ts.Columns(1).ColumnWidth = 6
    error_ts.Columns(2).ColumnWidth = 20
    error_ts.Columns(3).ColumnWidth = 10
    error_ts.Columns(4).ColumnWidth = 10
    error_ts.Columns(5).ColumnWidth = 70
    
    ' ���̑�����
    error_ts.Name = "���H���e�ꗗ"
    
    
    ' �D�揇�ɍ��킹�ĉ��H�������s��
    For order_count = 1 To order_max
        
        statusBar_text = "���̓f�[�^���H��ƒ�(" & Format(order_count) & "/" & order_max & ")"
        Application.StatusBar = statusBar_text
        
        ' �s����Ɠ��e���擾
        work_process = wb_process.Worksheets("���C�����j���[").Cells(31 + (order_count - 1), 31).Value
    
        ' ��Ɠ��e�ɍ��킹�Ċe���H�������R�[������
        Select Case work_process
            
            ' �t�Z�b�g����
            Case 1
            
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�t�Z�b�g����")
                Call Processing_Setreverse(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' ���{������
            Case 2
                
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("��+������")
                Call Processing_Complementarity(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' �r���I����
            Case 3
                
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�r���I����")
                Call Processing_Exclusive(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' �Z���N�g�t���O����
            Case 4
            
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�Z���N�g�t���O����")
                'Call Processing_Selectflg(ws_process, wsp_indata, indata_maxrow, statusBar_text)
                Call Processing_Selectflg3(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            ' �J�e�S���C�Y����
            Case 5
            
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�J�e�S���C�Y����")
                Call Processing_Categorize2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts, error_tb)
                
            Case 6
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�f�[�^�N���A����")
                Call data_clear1(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            Case 15
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("���~�b�g�}���`���H����")
                Call Processing_Limitmulti_2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            Case 16
                ' ��ƃV�[�g���Œ肵�������s���A�ݒ��ʏ�񂩂珈�����s����ws_process�͖��g�p
                Set ws_process = wb_process.Worksheets("�f�[�^�N���A����")
                Call data_clear2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)

            Case 17
                ' ��ƃV�[�g���Œ肵�������s��
                Set ws_process = wb_process.Worksheets("�������H����")
                Call amplification_data(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            Case 18
                ' �t�@�C������ύX���ۑ�
                
                ' ���H�������s���Ȃ������ꍇ
                If wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Value = "*���H��" Then
                
                    ' ���H��s���폜����
                    wsp_indata.Columns(wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column).Delete
                
                End If
                
                wb_indata.SaveAs _
                Filename:=file_path & "\1_DATA\" & ThisWorkbook.Worksheets("���C�����j���[").Range("H3").Value & "OT.xlsx"
                
            Case Else
    
        End Select
    
    Next order_count
        
'    Call Processing_Setreverse      ' �t�Z�b�g���H�p�v���V�[�W��
'    Call Processing_Exclusive       ' �r���I�������H�p�v���V�[�W��
'    Call Processing_Categorize      ' �J�e�S���C�Y�������H�p�v���V�[�W��
'    Call Processing_Complementarity ' ���{���������H�p�v���V�[�W��
    
    ' �X�e�[�^�X�o�[��������
    Application.StatusBar = False
    
    ' ��ʂւ̕\�����I���ɂ���
    Application.ScreenUpdating = True
    
    ' �J�e�S���C�Y�p���O�̍쐬
    'now_data = Now
    error_tb.SaveAs Filename:=file_path & "\4_LOG\" & Format(Now, "yyyymmddhhmmss") & _
    "_" & Mid(wb_process.Name, 1, Len(wb_process.Name) - 5) & "(���H���O).xlsx"
    error_tb.Close
    
    ' module04 �G���[�������̓��O�o�̓t�@�C�������
    'Close #30
    
    wb_process.Close SaveChanges:=False
    
    ' ���̓t�@�C�����N���[�Y����
    wb_indata.Close
    
    ' 0byte�t�@�C�����폜
    Call Finishing_Mcs2017
    Call Starting_Mcs2017
    
    ' �V�X�e�����O�̏o��
    ' 2020.6.4 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "05"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 05"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - ���̓f�[�^�̉��H�����F�g�p�t�@�C���m" & indata_logname & " �b " & Dir(filename_work) & "�n"
    Close #1
    
    
    ' ���b�Z�[�W���\���ɂ���
    Application.DisplayAlerts = True
    

    
    MsgBox "���̓f�[�^�̉��H���������܂����B", , "MCS2017"

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.19  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2018.06.14�@'
' �t�Z�b�g���H�p�v���V�[�W��                                                                       '
' �����P WorkSheet�^ �t�Z�b�g�����w���V�[�g                                                        '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Setreverse(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim qcode1_dataflg As Boolean   ' ��r�ݖ�(�q) ����t���O
    Dim qcode2_dataflg As Boolean   ' �t�Z�b�g�Ώېݖ�(�e)����t���O

    Dim ma_count As Long            ' MA�񓚓��e�m�F�p�J�E���g�ϐ�
    
    Dim force_setflg As Boolean     ' �����t�Z�b�g�t���O�i�[�p�ϐ�
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�t�Z�b�g���H������(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
        
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
        
        ' �Ώېݖ�̍s�ԍ����擾
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
        qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2))
        
        ' InputWord�̎擾
        input_word = ws_process.Cells(process_count, QCODE2_DATA1).Value
        
        ' �e��t���O��������
        process_flg = False
        force_setflg = False
        
        ' ����������s�� ���������s��Ȃ��ꍇ��"FALSE"
        If ws_process.Cells(process_count, QCODE2_DATA2).Value <> "" Then
            force_setflg = True
        End If
        
        ' ����������s�� ���������s��Ȃ��ꍇ��"FALSE"
        If ws_process.Cells(process_count, SKIP_FLG).Value = "" Then
            process_flg = True
        End If
        
        ' Input_Word���I�����͈͊O�̎��i�}���`�A���T�[�j
        If Val(input_word) > q_data(qcode2_row).ct_count And q_data(qcode2_row).q_format = "M" Then
            Call print_log("�t�Z�b�g����", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
            "���������ݐ��" & q_data(qcode2_row).q_code & "�Ɂu" & input_word & "�v�͑��݂��Ȃ����ߏ������s�킸�ɏI��", ws_logs)
            process_flg = False
        End If
        
        ' Input_Word���I�����͈͊O�̎��i���~�b�g�}���`�j
        If Val(input_word) > q_data(qcode2_row).ct_count And Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
            Call print_log("�t�Z�b�g����", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
            "���������ݐ��" & q_data(qcode2_row).q_code & "�Ɂu" & input_word & "�v�͑��݂��Ȃ����ߏ������s�킸�ɏI��", ws_logs)
            process_flg = False
        End If
        
        ' �������s���i�������肪�L�����Ή�����w�����S�čs���Ă��鎞�j
        If process_flg = True And qcode1_row <> 0 And qcode2_row <> 0 And input_word <> "" Then
        
            ' �N���A�������s���Ȃ����@�i�e�������񓚉j
            If force_setflg = True Then
            
                Call print_log("�t�Z�b�g����", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "�ɉ񓚂�����ꍇ�A" & q_data(qcode2_row).q_code & "�ɋ����I�Ɂu" _
                & input_word & "�v�����", ws_logs)
            
            ' �N���A�������s���Ȃ����@�i�e�������񓚉j
            ElseIf q_data(qcode1_row).q_format = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
            
                Call print_log("�t�Z�b�g����", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "�ɉ񓚂�����A" & q_data(qcode2_row).q_code & "�����񓚂̎��Ɂu" _
                & input_word & "�v�����", ws_logs)
            
            ' �N���A�������s����Ƃ��@�i�e���P��񓚁j
            Else
                Call print_log("�t�Z�b�g����", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "�ɉ񓚂�����A" & q_data(qcode2_row).q_code & "�����񓚂̎��Ɂu" _
                & input_word & "�v����́i" & q_data(qcode2_row).q_code & "�Ɂu" & input_word & "�v�ȊO�̉񓚂������" _
                & q_data(qcode1_row).q_code & "���N���A", ws_logs)
            End If
        
        
            ' ���̓f�[�^�S�Ă𔻒肷��
            For indata_count = START_ROW_INDATA To indata_maxrow
            
                ' �e��t���O��������
                qcode1_dataflg = False
                qcode2_dataflg = False
                
                ' �q�̐ݖ₪MA��������LM�̎�
                If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
                
                    ' �����񓚂̓��e���m�F���A�񓚂������qcode1_dataflg��L���ɂ���
                    For ma_count = 1 To q_data(qcode1_row).ct_count
                
                        ' �����ꂩ�̃J�e�S���[�ɉ񓚂�����ꍇ
                        If Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ma_count - 1))) <> "" Then
                            qcode1_dataflg = True
                        End If
                
                    Next ma_count
                
                Else
                    
                    ' ���̓f�[�^�̓��e�𔻒�
                    If Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value) <> "" Then
                
                        ' �������s��
                        qcode1_dataflg = True
                        
                    Else
                        
                        ' �������s��Ȃ�
                        qcode1_dataflg = False
                        
                    End If
                
                End If
                
                ' Qcode1�������\�̏ꍇ
                If qcode1_dataflg = True Then
                
                    ' �e�̐ݖ₪MA��������LM�̎�
                    If Mid(q_data(qcode2_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
                    
                        ' 20170519 0ct �Ή�
                        ' 0�J�e�S���[�t���O���L���̎�
                        'If q_data(qcode2_row).ct_0flg = True Then
                        
                            ' �t�Z�b�g�������s��
                        '    wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + Val(input_word)) = 1
                        
                        ' �ʏ폈��
                        'Else
                        
                            ' �t�Z�b�g�������s��
                            wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + (Val(input_word) - 1)) = 1
                        
                        'End If
                    
                    Else
                                            
                        ' �e�̉񓚓��e�𔻒�i�e�����񓚂̎��j
                        If Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) = "" Or _
                        ws_process.Cells(process_count, QCODE2_DATA2).Value <> "" Or force_setflg = True Then
                            
                            ' �t�Z�b�g�������s��
                            wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value = input_word
                        
                        ' �e�̉񓚓��e�𔻒�i�e�ɈقȂ�񓚂����鎞�j
                        ElseIf Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) <> Trim(input_word) Then
                        
                            ' �q�̐ݖ₪MA��������LM�̎�
                            If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
                                    
                                DoEvents
                                    
                                ' �q�̉񓚂�S�Ė��񓚏���
                                For ma_count = 1 To q_data(qcode1_row).ct_count
                
                                    wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ma_count - 1)) = ""
                
                                Next ma_count
                            
                            ' �q�̉񓚂����R�L�q�̎�
                            ElseIf q_data(qcode1_row).q_format = "F" Or q_data(qcode1_row).q_format = "O" Then
                                
                                ' �N���A�͍s��Ȃ�
                                
                            Else
                            
                                ' �q�̉񓚂𖳉񓚏���
                                wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value = ""
                        
                            End If
                        
                        ' �e�̉񓚓��e�𔻒�i�e�̉񓚂�����̎��j
                        ElseIf Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) <> input_word Then
                            
                            ' �������s��Ȃ�
                            
                        End If
                        
                    End If
                    
                End If
                
            Next indata_count
        
        End If
        
    Next process_count

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.19  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2018.06.14�@'
' �r���I�������H�p�v���V�[�W��                                                                     '
' �����P WorkSheet�^ �r���I�����w���V�[�g                                                          '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Exclusive(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
'    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
'    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim qcode1_dataflg As Boolean   ' ��r�ݖ�(�q) ����t���O
'    Dim qcode2_dataflg As Boolean   ' �t�Z�b�g�Ώېݖ�(�e)����t���O
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu

    Dim ma_count As Long            ' MA�񓚓��e�m�F�p�J�E���g�ϐ�
    
    Dim exclusive_ct As Long        ' �r���I�����Ώ۔ԍ��i�[�p�ϐ�
    
    Dim str_ct() As Variant         ' �J�e�S���[��r�p�z��i���̓f�[�^�j
    Dim str_maxcount As Long        ' �񓚐��i�[�p�ϐ�
    Dim str_min As Long             ' �ŏ��l�擾�p�ϐ�
    Dim str_address As String       ' MA�擪�A�h���X�i�[�p������ϐ�
    Dim str_target As Long          ' �z�����r�p�J�E���g�ϐ�
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�r���I������(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' QCODE���������ݒ��ʂ����ԍ����擾
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
    
        ' �w����MA��LM�̎�
        If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
            
            ' �r���I�J�e�S���[�̎w�����s���Ă���Ƃ����X�L�b�v�w�����Ȃ���
            If ws_process.Cells(process_count, QCODE1_DATA1).Value <> "" _
            And ws_process.Cells(process_count, SKIP_FLG).Value = "" Then
            
                ' �r���I�����J�e�S���[�ԍ����擾
                exclusive_ct = ws_process.Cells(process_count, QCODE1_DATA1).Value
                
                ' �J�e�S���[�ԍ����ݖ�͈͓̔����ǂ����𔻒�i�͈͓��j
                If q_data(qcode1_row).ct_count >= exclusive_ct And exclusive_ct <> 0 Then
                    
                    Call print_log("�r���I����", q_data(qcode1_row).q_code, "", q_data(qcode1_row).q_code & _
                    "�́u" & exclusive_ct & "�v�Ƒ��̉񓚂����݂��Ă��鎞�A�u" & exclusive_ct & "�v���N���A", ws_logs)
                    
                    ' ���̓f�[�^�ɏ������s��
                    For indata_count = START_ROW_INDATA To indata_maxrow
                    
                    
                        ' �Ώۂ̃J�e�S���[���L�����𔻒�
                        If wsp_indata.Cells(indata_count, _
                        q_data(qcode1_row).data_column + (exclusive_ct - 1)).Value = 1 Then
                   
                            ' MA�͈͂�S�Ĕz��Ɋi�[����
                            str_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            str_ct = wsp_indata.Range(str_address).Resize(, q_data(qcode1_row).ct_count)
                            ' �� str_ct( 1 to 1 , 1 to q_data(qcode1_row).ct_count )
                    
                            ' �z��ւ̉񓚐����擾
                            str_maxcount = Application.WorksheetFunction.Sum(str_ct)
                            str_min = Application.WorksheetFunction.Min(str_ct)
                            
                            ' �Ώۂ̃J�e�S���[�ȊO�ɂ��񓚂�����ꍇ
                            If str_maxcount > 1 Then
                                
                                ' ���񓚂֕ύX����
                                wsp_indata.Cells(indata_count, _
                                q_data(qcode1_row).data_column + (exclusive_ct - 1)).Value = ""
                        
                            End If
                    
                        End If
                    
                    Next indata_count
                
                ' �J�e�S���[�ԍ����ݖ�͈͓̔����ǂ����𔻒�i�͈͊O�j
                Else
                    Call print_log("�r���I����", q_data(qcode1_row).q_code, "", "��" & q_data(qcode1_row).q_code & _
                    "�̑I�����u" & exclusive_ct & "�v�͑��݂��Ă��Ȃ����߁A�������I�������܂����B", ws_logs)
                End If
                
            End If
            
        End If
        
    Next process_count
    
End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.20  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.04.21�@'
' ���{���������H�p�v���V�[�W��                                                                     '
' �����P WorkSheet�^ ���{�������w���V�[�g                                                          '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Complementarity(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu
    
    Dim str_ct1 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct2 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct3 As Variant          ' �J�e�S���[��r�p�z��i��+���j
    
    Dim str_maxcount As Long        ' �񓚐��i�[�p�ϐ�
    Dim str_address As String       ' MA�擪�A�h���X�i�[�p������ϐ�
    Dim str_target As Long          ' �z�����r�p�J�E���g�ϐ�
    
    Dim ct_count As Long            ' �z����e�J�E���g�p�ϐ�
    
    Dim ct1_count As Long           ' ���񓚐��i�[�p�ϐ�
    Dim ct2_count As Long           ' ���񓚐��i�[�p�ϐ�
    Dim ct3_count As Long           ' ��+���񓚐��i�[�p�ϐ�
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
        
        'Application.StatusBar = statusBar_text & _
        '"�@��+��������(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �X�L�b�v�t���O�������Ă��Ȃ������eQCODE���S�ċL������Ă��鎞
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE2) <> "" Then
        
        ' QCODE���������ݒ��ʂ����ԍ����擾
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
        qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2))
        
            ' ���w����MA��LM��S�̎�
            If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or _
            Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Or _
            Mid(q_data(qcode1_row).q_format, 1, 1) = "S" Then
                
                ' ���w����MA��LM�̎�
                If Mid(q_data(qcode2_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
        
                    ' �������e���m�F ������ and ������
                    If ws_process.Cells(process_count, QCODE2_DATA1) = "" Then
                        Call print_log("���{������", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & q_data(qcode1_row).q_code & _
                        "�i���j�̉񓚂�" & q_data(qcode2_row).q_code & "�i���j�ɒǉ����A�������񓚂����̉񓚐������̗L���񓚐��ȓ��̎��A���̉񓚓��e�����ɒǉ�", ws_logs)
                    ' �������e���m�F ������
                    Else
                        Call print_log("���{������", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & q_data(qcode1_row).q_code & _
                        "�i���j�̉񓚂�" & q_data(qcode2_row).q_code & "�i���j�ɒǉ�", ws_logs)
                    End If
                    
                    ' ���̓f�[�^�ɏ������s��
                    For indata_count = START_ROW_INDATA To indata_maxrow
        
                        ' ����T�C�Y�̔z����Ē�`
                        ReDim str_ct1(1, q_data(qcode1_row).ct_count)
                        ReDim str_ct2(1, q_data(qcode2_row).ct_count)
                        ReDim str_ct3(q_data(qcode1_row).ct_count)
        
                        ' ���w����S�̎�
                        If q_data(qcode1_row).q_format = "S" Then
                        
                            ' �񓚂���������
                            If wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <> "" Then
                            
                                ' �J�e�S���[�ʒu��1����
                                str_ct1(1, wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column)) = 1
                        
                            End If
                        
                        ' ���w����MA��LM�̎�
                        Else
                        
                            ' MA�͈͂�S�Ĕz��Ɋi�[����
                            str_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            str_ct1 = wsp_indata.Range(str_address).Resize(, q_data(qcode1_row).ct_count)
                        
                        End If
                        
                        ' ���w����S�Ĕz��Ɋi�[����
                        str_address = wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Address
                        str_ct2 = wsp_indata.Range(str_address).Resize(, q_data(qcode2_row).ct_count)
                        
                        ' �ő�J�e�S���[�����擾�i���̕����J�e�S���[���傫������str_ct2���g�p�j
                        ' 20180618 �������J�e�S���[���ɍ��킹�z���p��
                        If q_data(qcode1_row).ct_count = q_data(qcode2_row).ct_count Or _
                        q_data(qcode1_row).ct_count > q_data(qcode2_row).ct_count Then
                            str_maxcount = q_data(qcode2_row).ct_count
                        ElseIf q_data(qcode1_row).ct_count < q_data(qcode2_row).ct_count Then
                            str_maxcount = q_data(qcode1_row).ct_count
                        End If
                        

                        
                        ' ��(str_ct1)�̉񓚏󋵂Ɓ�(str_ct2)�̉񓚏󋵂����킹str_ct3���쐬����
                        For ct_count = 1 To str_maxcount
                        
                            'str_ct3(ct_count) = str_ct1(1, ct_count) Or str_ct2(1, ct_count)
                            ' ���������́��ŉ񓚂�����ꍇ
                            If str_ct1(1, ct_count) > 0 Or str_ct2(1, ct_count) > 0 Then
                            
                                ' str_ct3�ɂ��킹���f�[�^���쐬����
                                ' 20180618 ���H�i�K�ł͉񓚂̓��e��킸1���Ă̂�
                                str_ct3(ct_count) = 1
                            
                            Else
'                               str_ct3(ct_count) = ""
                            End If
                        
                        Next ct_count
                        
                        ' �e�z��̉񓚐����i�[
                        ct1_count = Application.WorksheetFunction.Sum(str_ct1)
                        ct2_count = Application.WorksheetFunction.Sum(str_ct2)
                        ct3_count = Application.WorksheetFunction.Sum(str_ct3)
                        
                        DoEvents
                        
                        ' ���Z���ĉ񓚐��������Ă���ꍇ
                        If ct2_count < ct3_count Then
                            
'                            wsp_indata.Range(str_address).Resize(, q_data(qcode2_row).ct_count) = str_ct3

                            ' str_ct3�̓��e����(str_ct2)�֏㏑������
                            For ct_count = 1 To str_maxcount
                            
                                ' str_ct3�ɂ��킹���f�[�^���쐬����
                                wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + (ct_count - 1)) _
                                = str_ct3(ct_count)
 
                            Next ct_count
                            
                            ' �J�e�S���[�������킹�ĕύX
                            ' 20180618 ���Z�͈͂�ύX�������ߍ폜
                            'ct2_count = ct3_count
                            
                        End If
                        
                        ' �Z�b�g���s����
                        If ws_process.Cells(indata_count, QCODE2_DATA1).Value = "" Then
                        
                            ' ����SA�����̉񓚐���0�A���̉񓚐���1�̎�
                            If q_data(qcode1_row).q_format = "S" And ct1_count = 0 And ct2_count = 1 Then
                        
                                ' str_ct2�̓��e����(str_ct1)�֏㏑������
                                For ct_count = 1 To str_maxcount
                            
                                    ' ���̉񓚓��e�Ɠ����J�e�S���[��L���ɂ���
                                    If str_ct2(1, ct_count) > 0 Then
                                
                                        ' �������֒ǉ�����
                                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value = ct_count
                                        '
                                        Exit For
                                
                                    End If
                            
                                Next ct_count
                        
                            ' �������񓚂����̉񓚐������̃��[�v�J�E���g�ȉ��̎�
                            ElseIf ct1_count = 0 And q_data(qcode1_row).ct_loop >= ct2_count Then
                        
                                ' �ő�J�e�S���[�����擾�i���̃J�e�S���[���ɍ��킹���str_ct1���g�p�j
                                str_maxcount = q_data(qcode2_row).ct_count
                        
                                ' str_ct2�̓��e����(str_ct1)�֏㏑������
                                For ct_count = 1 To str_maxcount
                            
                                    ' str_ct3�ɂ��킹���f�[�^���쐬����
                                    wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ct_count - 1)) _
                                    = str_ct2(1, ct_count)
 
                                Next ct_count
                        
                            End If
                        
                        End If
                        
                    Next indata_count
        
                ' ���w���������񓚂ł͂Ȃ���
                Else
                
                    ' �G���[�R�����g�̏o��
                    Call print_log("���{������", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & _
                    q_data(qcode2_row).q_code & "�i���j�̌`���������񓚐ݖ�ł͂���܂���", ws_logs)
                
                End If
                
            End If

        End If
    
    Next process_count


End Sub

'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2017.05.08  '
' �Z���N�g�t���O���H�p�v���V�[�W��                                                                 '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim qcode3_row As Long          ' �G���g���[�G���A�I�[�i�[�p�ϐ�
    
    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu
    
    Dim str_ct1 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct2 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct3 As Variant          ' �J�e�S���[��r�p�z��i��+���j
    
    Dim str_maxcount As Long        ' �񓚐��i�[�p�ϐ�
    Dim str_address As String       ' MA�擪�A�h���X�i�[�p������ϐ�
    Dim str_target As Long          ' �z�����r�p�J�E���g�ϐ�
    
    Dim ct_count As Long            ' �z����e�J�E���g�p�ϐ�
    
'    Dim ct1_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct2_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct3_count As Long           ' ��+���񓚐��i�[�p�ϐ�
    
    Dim work1_flg As Boolean        ' �������i�[�p�t���O
    Dim work2_flg As Boolean        ' �������i�[�p�t���O
    
    Dim column_max As Long          ' AND�AOR�����I�[�ԍ��i�[�p�ϐ�
    Dim and_column() As Long        ' ���������Ή��p���I�z��
    Dim and_data() As Variant       ' �����������i�[�p�ϐ�
    Dim and_count As Long           ' AND�AOR����
    Dim and_target As Long          ' �Ώۗ�ԍ��i�[�p�ϐ�
    
    Dim processing_flg As Boolean   ' Function�߂�l�i�[�p�ϐ�
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' ���������̏I�[�ʒu���擾
    column_max = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' �A�h���X�ʒu�i�[�p�z����Ē�`
    ReDim and_column(300)
    
    ' �Z���N�g�����̏����ʒu��ݒ�
    and_count = 1
    and_column(1) = SF_START
    
    ' �A�h���X�ʒu���擾
    For and_target = QCODE1_DATA6 To column_max
                    
        ' �����̍s�𔻒�
        If ws_process.Cells(START_ROW - 1, and_target).Value = "�ڑ���" & vbLf & "�i���������j" Then
        
            ' �������݈ʒu��ύX
            and_count = and_count + 1
            
            ' �J��������z��Ɋi�[
            and_column(and_count) = and_target
            
        End If
        
    Next
    
    ' �z����Ē�`�i�l�͕ێ�����j
    ReDim Preserve and_column(and_count)
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�Z���N�g���H������(" & Format(process_count - START_SUTATUSBER) & _
        '"/" & Format(process_max - START_SUTATUSBER) & ")"

        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �X�L�b�v�t���O�������Ă��Ȃ������eQCODE���S�ċL������Ă��鎞
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA3) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA4) <> "" Then
        
            ' QCODE���}�b�`���O�������ԍ����i�[
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4))
            qcode3_row = Qcode_Match("*���H��")
        
            ' ���ɏo�̓G���A���p�ӂ���Ă��鎞�i�����H�������ɏo�̓G���A���ݒ肳��Ă��鎞
            If q_data(qcode2_row).data_column <> 0 And _
            q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' �ʏ폈��
                If q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' EntryArea�ɏ������ގw���̏ꍇ
                Else
                                        
                    'Print #30, "���H�悪EntryArea�ɐݒ肳��Ă��邽�ߏ������s���܂���ł����A�m�F�����肢���܂��B"
            
                End If
            
            ' �܂��G���A���p�ӂ���Ă��Ȃ���
            Else
        
                ' �w�b�_���쐬����
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' �V�����ݒ肵���G���A��q_data�ɃJ�����Ƃ��Đݒ肷��
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' �ŏ��l�A�ő�l����̓f�[�^�̃w�b�_�Ɋi�[
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' ���H�w���̓��e��񂲂Ǝ擾
            and_data = ws_process.Range(ws_process.Cells(process_count, 1).Address).Resize(, column_max - 1)
            'ReDim Preserve and_data(column_max - 1)
            
            ' �Z���N�g�t���O�̏����ɍ��킹�ĉ񓚓��e�𔻒肷��
            Call salectflg_decision(wsp_indata, and_data, and_column, indata_maxrow, qcode2_row)
            'Call data_clear(wsp_indata, and_)
        
        
        ' �z��I�[�܂Ŕ���
        'For column_count = 1 To UBound(and_column)
        'Next
    
        End If
    
    Next process_count
    
End Sub


'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.21  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.02�@'
' �J�e�S���C�Y�������H�p�v���V�[�W��                                                               '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Categorize(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim str_address() As String     ' �����A�h���X�i�[�p�z��
    Dim str_qcode() As Long         ' ����QCODE�i�[�p�z��
    Dim target_coderow As Long      ' �Ώ�QCODE��ԍ��ꎞ�i�[�p�ϐ�
    
    Dim column_count As Long        ' ������ԍ��i�[�p�ϐ�
    Dim column_end As Long          ' ��I�[�ԍ��i�[�p�ϐ�
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�

    Dim str_ctdata As Variant       ' �J�e�S���C�Y�e�[�u�����i�[�p�z��ϐ�
    Dim str_ctdata_c As Range       ' �ꎞ�ۑ��p�ϐ�
    
    Dim writing_column As Long      ' �������݈ʒu�i�[�p�ϐ�
    Dim target_data As Double       ' ���̓f�[�^����p�ϐ�
    Dim categorize_count As Long    ' ����J�E���g�p�ϐ�
    Dim table_max As Long           ' �I�[�e�[�u���ԍ��i�[�p�ϐ�
    Dim ma_count As Long            ' MA�J�e�S���[�����擾
    
    ' ���̓f�[�^MA����p�ϐ��Q
    Dim str_madata As Variant       ' MA�񓚃f�[�^�i�[�p�ϐ�
    Dim ma_address As String        ' MA�񓚃f�[�^�A�h���X�i�[�p�ϐ�
    Dim maindata_count As Long      ' MA�f�[�^�J�E���g�ϐ�
    
    Dim prosessing_flg As Boolean   ' �w�b�_���H����p�t���O

    Dim pmax_number As Double       ' �e�[�u���ő�l�i�[�p�ϐ�
    Dim pmin_number As Double       ' �e�[�u���ŏ��l�i�[�p�ϐ�

    Dim str_count As Long           ' str_ctdata�J�E���g�p�ϐ�

    ' �����ݒ�
    ReDim str_address(300)
    ReDim str_qcode(300)
    
    ' ��ʂւ̕\�����I���ɂ���
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "�@�J�e�S���C�Y���H�������v�Z��..."
    
    ' ��ʂւ̕\�����I�t�ɂ���
    'Application.ScreenUpdating = False
    
    ' �I�[��ԍ��̎擾 20170502 START_ROW-1��5�ɕύX
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column

    ' �����������擾
    For column_count = 11 To column_end
    
        ' �A�X�^���X�N���������Ƃ����𐔂���
        ' ���A�X�^���X�N�Ɗ����b�s���Œ�ʒu���Ώېݖ�ɋL�����聕Skip�t���O����
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 1, column_count + 5).Value = "�����b�s" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 5).Value = "" Then
            
            ' QCODE���擾
            target_coderow = Qcode_Match(ws_process.Cells(START_ROW - 2, column_count + 2).Value)
            
            ' ���e�����l�̎�
            If IsNumeric(target_coderow) Then
            
                ' ������Ώی����𑝂₷
                process_max = process_max + 1
            
                ' QCODE��ԍ���z��Ɋi�[����
                str_qcode(process_max) = target_coderow
            
                ' �A�X�^���X�N�̈ʒu��z��Ɋi�[����
                str_address(process_max) = ws_process.Cells(START_ROW - 1, column_count).Address

            End If
            
        End If
    
    Next column_count
    
    ' �i�[���������ɉ����Ĕz����Ē�`�i�l�͕ێ�����j
    ReDim Preserve str_address(process_max)
    
    ' �J�e�S���C�Y�������s���i�z��Ɏ�荞�񂾎w�������j
    For process_count = 1 To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�J�e�S���C�Y���H������(" & Format(process_count) & "/" & Format(process_max) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �e�[�u������z��Ɋi�[
        Set str_ctdata_c = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
        str_ctdata = str_ctdata_c.NumberFormatLocal
        
        ' �������݈ʒu���擾
        writing_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        
        ' ���̓f�[�^�G���A��QCODE��ݒ�
        ' wsp_indata.Cells(1, writing_column) = ws_process.Range(str_address(process_count)).Offset(-1, 3).Value
        
        ' �w�b�_�����쐬
        prosessing_flg = Hedder_Create(wsp_indata, ws_process.Range(str_address(process_count)) _
        .Offset(-1, 3).Value, wsp_indata.Cells(1, writing_column).Address)
        
        ' �l�`�o�͏����擾
        ma_count = Val(ws_process.Range(str_address(process_count)).Offset(-1, 4))
        
        ' �l�`�ȊO�ŏo�͂��s����
        If ma_count = 0 Then
            
            ' �����l�ݒ�
            pmax_number = Val(str_ctdata(1, 6))
            pmin_number = Val(str_ctdata(1, 6))
            
            ' �w�b�_�[�쐬�p
            For str_count = 1 To 300
            
                ' �ő�l�A�ŏ��l���擾
                If str_ctdata(str_count, 6) <> "" Then
                
                    ' �ő�l
                    If pmax_number < Val(str_ctdata(str_count, 6)) Then
                        pmax_number = Val(str_ctdata(str_count, 6))
                    End If
                    
                    ' �ŏ��l
                    If pmin_number > Val(str_ctdata(str_count, 6)) Then
                        pmin_number = Val(str_ctdata(str_count, 6))
                    End If
                
                Else
                    Exit For
                End If
            
            Next
            
            ' �ŏ��l�ƍő�l����������
            wsp_indata.Cells(5, writing_column).Value = pmin_number
            wsp_indata.Cells(6, writing_column).Value = pmax_number
        
        End If
        
        ' �I�[�e�[�u���ԍ����擾
        table_max = UBound(str_ctdata)
        
        ' ���̓f�[�^�S�ĂɃJ�e�S���C�Y�������s��
        For indata_count = START_ROW_INDATA To indata_maxrow
            
            ' ���̉񓚋敪�ɂ���ď������킯��
            Select Case q_data(str_qcode(process_count)).q_format

                ' SA�������͎����񓚌n��
                Case "S", "R", "H"
            
                    ' ���͒l�����݂���ꍇ�������s��
                    If wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value <> "" Then
                        
                        ' ���̓f�[�^���擾
                        target_data = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
            
                        ' �e�[�u�����̏������܂Ƃ߂čs��
                        For categorize_count = 1 To table_max
                    
                            ' �J�e�S���C�Y�w���������ꍇ
                            If str_ctdata(categorize_count, 3) = "" Then
                                Exit For
                            End If
                    
                            ' �A�X�^���X�N�ӏ��ɏ�񂪋L������Ă��Ȃ����������s��
                            If str_ctdata(categorize_count, 1) = "" Then
                        
                                ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                If target_data >= str_ctdata(categorize_count, 3) And target_data <= _
                                str_ctdata(categorize_count, 5) Then
                            
                                    ' MA�o�͎w���𔻒�(SA�o��)
                                    If ma_count = 0 Then
                                        
                                        ' ����CT���o�͐�ɐݒ�
                                        wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 6)
                                        
                                    ' MA�o�͎w���𔻒� (MA�o��)
                                    Else
                                        
                                        ' ����CT���o�͐�ɐݒ�
                                        wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 6) - 1)) = 1
                            
                                    End If
                        
                                ' �e�[�u���͈͊O�̋L���l�̏ꍇ
                                Else
                                
                                    ' �����O���o�͂���
                                
                                End If
                        
                            End If
            
                        Next categorize_count
                    End If
                
                ' MA
                Case "M"
                    
                    ' ���̓f�[�^�̐擪�A�h���X���擾
                    ma_address = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Address
                    
                    ' MA�񓚔͈͂�z��Ƃ��Ď擾
                    str_madata = wsp_indata.Range(ma_address).Resize(, (q_data(str_qcode(process_count)).ct_count))
                    'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
                                        
                    ' �񓚂����݂��Ă���ꍇ
                    If WorksheetFunction.Sum(str_madata) <> 0 Then
                    
                        ' MA�����J�e�S���C�Y
                        For maindata_count = 1 To q_data(str_qcode(process_count)).ct_count
                        
                            ' ���̓f�[�^�����݂��Ă���ꍇ�i�O�ȏ�̐��l�j
                            If str_madata(1, maindata_count) > 0 Then
            
                                ' �e�[�u�����̏������܂Ƃ߂čs��
                                For categorize_count = 1 To table_max
                    
                                    ' �J�e�S���C�Y�w���������ꍇ
                                    If str_ctdata(categorize_count, 3) = "" Then
                                        Exit For
                                    End If
                    
                                    ' �A�X�^���X�N�ӏ��ɏ�񂪋L������Ă��Ȃ����������s��
                                    If str_ctdata(categorize_count, 1) = "" Then
                        
                                        ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                        If maindata_count >= str_ctdata(categorize_count, 3) And maindata_count <= _
                                        str_ctdata(categorize_count, 5) Then
                            
                                            ' ����CT���o�͐�ɐݒ�
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 6) - 1)) = 1
                        
                                        
                                        ' �e�[�u���͈̔͊O�̋L���l�̏ꍇ
                                        
                                            ' �����O���o�͂���
                                        
                                        End If
                        
                                    End If
            
                                Next categorize_count
                            
                            End If
                        
                        Next maindata_count
                    
                    End If
   
                Case Else
                    
            End Select

        Next indata_count
    
    Next process_count

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.21  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.02�@'
' �J�e�S���C�Y�������H�p�Q�v���V�[�W��                                                             '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
' �����T Workbook�^  ���H���O�o�̓u�b�N                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Categorize2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet, ByVal error_tb As Workbook)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim str_address() As String     ' �����A�h���X�i�[�p�z��
    Dim str_qcode() As Long         ' ����QCODE�i�[�p�z��
    Dim str_outqcode() As String    ' �o��QCODE�i�[�p�z��
    Dim target_coderow As Long      ' �Ώ�QCODE��ԍ��ꎞ�i�[�p�ϐ�
    
    Dim column_count As Long        ' ������ԍ��i�[�p�ϐ�
    Dim column_end As Long          ' ��I�[�ԍ��i�[�p�ϐ�
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�

    Dim str_ctdata As Variant       ' �J�e�S���C�Y�e�[�u�����i�[�p�z��ϐ�
    Dim str_ctpro As String         ' �J�e�S���C�Y�e�[�u�������ύX�p�ϐ�
    
    Dim writing_column As Long      ' �������݈ʒu�i�[�p�ϐ�
    Dim target_data As Double       ' ���̓f�[�^����p�ϐ�
    Dim categorize_count As Long    ' ����J�E���g�p�ϐ�
    Dim table_max As Long           ' �I�[�e�[�u���ԍ��i�[�p�ϐ�
    Dim ma_count As Long            ' MA�J�e�S���[�����擾
    
    ' ���̓f�[�^MA����p�ϐ��Q
    Dim str_madata As Variant       ' MA�񓚃f�[�^�i�[�p�ϐ�
    Dim ma_address As String        ' MA�񓚃f�[�^�A�h���X�i�[�p�ϐ�
    Dim maindata_count As Long      ' MA�f�[�^�J�E���g�ϐ�
    
    Dim prosessing_flg As Boolean   ' �w�b�_���H����p�t���O

    Dim pmax_number As Double       ' �e�[�u���ő�l�i�[�p�ϐ�
    Dim pmin_number As Double       ' �e�[�u���ŏ��l�i�[�p�ϐ�

    Dim str_count As Long           ' str_ctdata�J�E���g�p�ϐ�
    
    Dim identity_count As Long      ' �����NO�J�E���g�p�ϐ�
    Dim identity_max As Long        ' �����NO�ő吔�i�[�p�ϐ�
    
    Dim stray_ws As Worksheet       ' ���J�e�S���C�Y�f�[�^���O�o�͗p�I�u�W�F�N�g�ϐ�
    Dim ct_flg As Boolean           ' ���J�e�S���C�Y�f�[�^����p�t���O
    Dim stray_row As Long           ' ���J�e�S���C�Y�f�[�^���O�o�̓A�h���X�p�ϐ�
    
    'Dim log_rows As Long            ' ���O�o�͈ʒu�i�[�p�ϐ�
    
    ' �����ݒ�
    ReDim str_address(300)
    ReDim str_qcode(300)
    ReDim str_outqcode(300)
    
    ' ��ʂւ̕\�����I���ɂ���
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "�@�J�e�S���C�Y���H�������v�Z��..."
    
    ' ��ʂւ̕\�����I�t�ɂ���
    'Application.ScreenUpdating = False
    
    
    ' �I�[��ԍ��̎擾 20170502 START_ROW ��6�ɕύX
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' ���J�e�S���C�Y���i�[�p�V�[�g�ǉ�
    Set stray_ws = error_tb.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    stray_row = 2
    
    stray_ws.Name = "���J�e�S���C�Y���X�g"
    stray_ws.Range("A1").Value = "SampleNo"
    stray_ws.Range("B1").Value = "QCODE"
    stray_ws.Range("C1").Value = "MA_CT"
    stray_ws.Range("D1").Value = "�G���[���e"
    stray_ws.Range("E1").Value = "�񓚓��e"
    stray_ws.Range("F1").Value = "�C�����e"
    'stray_ws.Range("G1").Value = "����"
    
    stray_ws.Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlHAlignCenter
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(58, 56, 56)
    End With
    stray_ws.Rows(1).RowHeight = 18
    stray_ws.Range("A:C").EntireColumn.ColumnWidth = 8.5
    stray_ws.Columns("D:D").ColumnWidth = 49.88
    stray_ws.Range("E:F").EntireColumn.ColumnWidth = 8.88

    ' �����������擾
    For column_count = 12 To column_end
    
        ' �A�X�^���X�N���������Ƃ����𐔂���
        ' ���A�X�^���X�N�Ɗ����b�s���Œ�ʒu���Ώېݖ�ɋL�����聕Skip�t���O����
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 1, column_count + 8).Value = "�����b�s" And _
        ws_process.Cells(START_ROW, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value = "" Then
        
            
            identity_max = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row - 5
            
            ' �����m�n�����[�v
            For identity_count = 1 To identity_max
            
                ' QCODE���擾
                If ws_process.Cells(START_ROW + (identity_count - 1), column_count).Value = "" Then
                    target_coderow = Qcode_Match(ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value)
                Else
                    target_coderow = 0
                End If
            
                ' ���e�����l�̎�
                If target_coderow <> 0 Then
            
                    ' ������Ώی����𑝂₷
                    process_max = process_max + 1
            
                    ' QCODE��ԍ���z��Ɋi�[����
                    str_qcode(process_max) = target_coderow
                    str_outqcode(process_max) = ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value
            
                    ' �A�X�^���X�N�̈ʒu��z��Ɋi�[����
                    str_address(process_max) = ws_process.Cells(START_ROW - 1, column_count).Address
                    
                    ' �������e���o��
                    Call print_log("�J�e�S���C�Y����", ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value, _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value, _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value & "�̉񓚂�" & _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value & "�փJ�e�S���C�Y�o�́B", ws_logs)

                End If
                
            Next identity_count
        
        End If
    
    Next column_count
    
    ' �i�[���������ɉ����Ĕz����Ē�`�i�l�͕ێ�����j
    ReDim Preserve str_address(process_max)
    
    ' �J�e�S���C�Y�������s���i�z��Ɏ�荞�񂾎w�������j
    For process_count = 1 To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�J�e�S���C�Y���H������(" & Format(process_count) & "/" & Format(process_max) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �e�[�u������z��Ɋi�[
        
        str_ctpro = ws_process.Range(str_address(process_count)).Offset(1, 5).Resize(300, 4).Address
        str_ctdata = ws_process.Range(str_ctpro).Value2
        'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 5).Resize(300, 4).Value
        
        ' �������݈ʒu���擾
        writing_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        
        ' ���̓f�[�^�G���A��QCODE��ݒ�
        ' wsp_indata.Cells(1, writing_column) = ws_process.Range(str_address(process_count)).Offset(-1, 3).Value
        
        ' �w�b�_�����쐬
        'prosessing_flg = Hedder_Create(wsp_indata, q_data(process_count).q_code _
        ', wsp_indata.Cells(1, writing_column).Address)
        
        ' �w�b�_�����쐬
        prosessing_flg = Hedder_Create(wsp_indata, str_outqcode(process_count), _
        wsp_indata.Cells(1, writing_column).Address)
        
        ' �l�`�o�͏����擾
        ma_count = Val(ws_process.Range(str_address(process_count)).Offset(-1, 3))
        
        ' �l�`�ȊO�ŏo�͂��s����
        If ma_count = 0 Then
            
            ' �����l�ݒ�
            pmax_number = Val(str_ctdata(1, 4))
            pmin_number = Val(str_ctdata(1, 4))
            
            ' �w�b�_�[�쐬�p
            For str_count = 1 To 300
            
                ' �ő�l�A�ŏ��l���擾
                If str_ctdata(str_count, 4) <> "" Then
                
                    ' �ő�l
                    If pmax_number < Val(str_ctdata(str_count, 4)) Then
                        pmax_number = Val(str_ctdata(str_count, 4))
                    End If
                    
                    ' �ŏ��l
                    If pmin_number > Val(str_ctdata(str_count, 4)) Then
                        pmin_number = Val(str_ctdata(str_count, 4))
                    End If
                
                Else
                    Exit For
                End If
            
            Next
            
            ' �ŏ��l�ƍő�l����������
            wsp_indata.Cells(5, writing_column).Value = pmin_number
            wsp_indata.Cells(6, writing_column).Value = pmax_number
        
        End If
        
        ' �I�[�e�[�u���ԍ����擾
        table_max = UBound(str_ctdata)
        
        
        
        ' ���̓f�[�^�S�ĂɃJ�e�S���C�Y�������s��
        For indata_count = START_ROW_INDATA To indata_maxrow
            
            ' ���̉񓚋敪�ɂ���ď������킯��
            Select Case q_data(str_qcode(process_count)).q_format

                ' SA�������͎����񓚌n��
                Case "S", "R", "H"
            
                    ' ���͒l�����݂���ꍇ�������s��
                    If wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value <> "" Then
                        
                        ' ���J�e�S���C�Y���X�g�o�͗p�t���O�i���J�e�S���C�Y�� False �j
                        ct_flg = False
                        
                        ' ���̓f�[�^���擾
                        target_data = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
            
                        ' �e�[�u�����̏������܂Ƃ߂čs��
                        For categorize_count = 1 To table_max
                    
                            ' �J�e�S���C�Y�w���������ꍇ
                            If str_ctdata(categorize_count, 1) = "" Then
                                Exit For
                            End If
                    
                            ' �A�X�^���X�N�ӏ��ɏ�񂪋L������Ă��Ȃ����������s��
                            'If str_ctdata(categorize_count, 1) = "" Then
                            
                            ' �͈͎w�肪����ꍇ
                            If str_ctdata(categorize_count, 3) <> "" Then
                            
                                ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                If target_data >= str_ctdata(categorize_count, 1) And target_data <= _
                                str_ctdata(categorize_count, 3) Then
                            
                                    ' MA�o�͎w���𔻒�(SA�o��)
                                    If ma_count = 0 Then
                                        
                                        ' ����CT���o�͐�ɐݒ�
                                        wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                        ct_flg = True
                                        Exit For
                                        
                                    ' MA�o�͎w���𔻒� (MA�o��)
                                    Else
                                        
                                        ' ����CT���o�͐�ɐݒ�
                                        wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                        ct_flg = True
                                        Exit For
                            
                                    End If
                                
                                ' ����Y��(���J�e�S���C�Y�j
                                Else
                                
                                End If
                            
                            ' �n�_�w���݂̂̏ꍇ
                            Else
                                ' ���̎n�_�����鎞
                                If str_ctdata(categorize_count + 1, 1) <> "" Then
                            
                                    ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                    If target_data >= str_ctdata(categorize_count, 1) And target_data < _
                                    str_ctdata(categorize_count + 1, 1) Then
                            
                                        ' MA�o�͎w���𔻒�(SA�o��)
                                        If ma_count = 0 Then
                                        
                                            ' ����CT���o�͐�ɐݒ�
                                            wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                            ct_flg = True
                                            Exit For
                                        
                                        ' MA�o�͎w���𔻒� (MA�o��)
                                        Else
                                        
                                            ' ����CT���o�͐�ɐݒ�
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                            ct_flg = True
                                            Exit For
                            
                                        End If
                                    
                                    ' ���J�e�S���C�Y
                                    Else
                                        
                                    End If
                                    
                                ' �n�_���Ȃ���
                                Else
                                
                                    ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                    If target_data >= str_ctdata(categorize_count, 1) Then
                            
                                        ' MA�o�͎w���𔻒�(SA�o��)
                                        If ma_count = 0 Then
                                        
                                            ' ����CT���o�͐�ɐݒ�
                                            wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                            ct_flg = True
                                            Exit For
                                        
                                        ' MA�o�͎w���𔻒� (MA�o��)
                                        Else
                                        
                                            ' ����CT���o�͐�ɐݒ�
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                            ct_flg = True
                                            Exit For
                            
                                        End If
                                    
                                    ' ���J�e�S���C�Y
                                    Else
                                        
                                    
                                    End If
                                
                                End If
                            
                            End If
                            
                            'End If
            
                        Next categorize_count
                        
                        ' �J�e�S���C�Y���s��Ȃ������ꍇ�A���O���o��
                        If ct_flg = False Then

                            ' SampleNo
                            stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
                            ' QCODE
                            stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
                            ' MA_CT
                            'stray_ws.Cells(stray_row, 3).Value = 1
                            ' �G���[���e
                            stray_ws.Cells(stray_row, 4).Value = "�uTable�ԍ��@" & Format(ws_process.Range(str_address(process_count)).Offset(-3, 3), "000") & "�v�@���J�e�S���C�Y"
                            ' �񓚓��e
                            stray_ws.Cells(stray_row, 5).Value = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
               
                            stray_row = stray_row + 1
                            
                        End If
                    
                    ' ���͒l�����݂��Ȃ��ꍇ
                    Else
                    
                    End If
                
                ' MA
                Case "M"
                    
                    ' ���̓f�[�^�̐擪�A�h���X���擾
                    ma_address = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Address
                    
                    ' MA�񓚔͈͂�z��Ƃ��Ď擾
                    str_madata = wsp_indata.Range(ma_address).Resize(, (q_data(str_qcode(process_count)).ct_count))
                    'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
                                        
                    ' �񓚂����݂��Ă���ꍇ
                    If WorksheetFunction.Sum(str_madata) <> 0 Then
                    
                        ' MA�����J�e�S���C�Y
                        For maindata_count = 1 To q_data(str_qcode(process_count)).ct_count
                        
                            ' ���̓f�[�^�����݂��Ă���ꍇ
                            If str_madata(1, maindata_count) > 0 Then
            
                                ' �e�[�u�����̏������܂Ƃ߂čs��
                                For categorize_count = 1 To table_max
                    
                                    ' �J�e�S���C�Y�w���������ꍇ
                                    If str_ctdata(categorize_count, 1) = "" Then
                                        Exit For
                                    End If
                    
                                    ' �A�X�^���X�N�ӏ��ɏ�񂪋L������Ă��Ȃ����������s��
                                    'If str_ctdata(categorize_count, 1) = "" Then
                        
                                        ' �e�[�u���͈͓̔��̋L���l�̏ꍇ
                                        If maindata_count >= str_ctdata(categorize_count, 1) And maindata_count <= _
                                        str_ctdata(categorize_count, 3) Then
                            
                                        ' 20180629 �o�͐�̌`�Ԃɍ��킹�ďC���ɕύX
                                        ' �J�e�S���C�Y�̃e�[�u�����t�ɗp�ӂ��邱�Ƃłl�`�̃V���O�����Ɏg���邽��
                                            
                                            ' MA�ł̏o��
                                            If ma_count <> 0 Then
                                            
                                                ' ����CT���o�͐�ɐݒ�
                                                wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                                ct_flg = True
                                                Exit For
                                            
                                            ' MA�ȊO�ł̏o��
                                            Else
                                            
                                                ' ����CT���o�͐�ɐݒ�
                                                wsp_indata.Cells(indata_count, writing_column) = maindata_count
                                                ct_flg = True
                                                Exit For
                                            
                                            End If
                                        
                                        ' ����Y��(���J�e�S���C�Y�j�e�[�u���͈͊O�̋L���l�̏ꍇ
                                        Else
                                        
                                            ' SampleNo
                                            stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
                                            ' QCODE
                                            stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
                                            ' MA_CT
                                            stray_ws.Cells(stray_row, 3).Value = maindata_count
                                            ' �G���[���e
                                            stray_ws.Cells(stray_row, 4).Value = "�uTable�ԍ��@" & Format(ws_process.Range(str_address(process_count)).Offset(-3, 3), "000") & "�v�@���J�e�S���C�Y"
                                            ' �񓚓��e
                                            stray_ws.Cells(stray_row, 5).Value = wsp_indata.Cells(indata_count, (q_data(str_qcode(process_count)).data_column + maindata_count - 1)).Value
               
                                            stray_row = stray_row + 1
                                        
                                        End If
                                    
                                    'End If
                                
                                Next categorize_count
                            
                            ' ���͒l�����݂��Ȃ��ꍇ
                            Else
                            
                            End If
                        
                        Next maindata_count
                    
                    End If
   
                Case Else
                    
            End Select
            
            ' �J�e�S���C�Y�ł��Ȃ��������R�[�h�����o��
            'If ct_flg = False Then
            '
            '    ' SampleNo
            '    stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
            '    ' QCODE
            '    stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
            '    ' MA_CT
            '    'stray_ws.Cells(stray_row, 3).Value = 1
            '    ' �G���[���e
            '    stray_ws.Cells(stray_row, 4).Value = "���J�e�S���C�Y"
            '    ' �񓚓��e
            '    stray_ws.Cells(stray_row, 5).Value = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
            '    'stray_ws.Cells(stray_row, 0).Value = 1
            '    'stray_ws.Cells(stray_row, 0).Value = 1
            '
            '    stray_row = stray_row + 1
            '
            'End If
        
        Next indata_count
    
    Next process_count
    


End Sub



'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.26  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.04.26�@'
' ���~�b�g�}���`���H�����p�v���V�[�W��                                                             '
' �����P WorkSheet�^ ���~�b�g�}���`���H�����w���V�[�g                                              '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Limitmulti_1(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
'    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
'    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
'    Dim qcode2_dataflg As Boolean   ' �t�Z�b�g�Ώېݖ�(�e)����t���O
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu

    Dim process_case As Long        ' �������e�i�[�p�ϐ�
    Dim start_address As String     ' �k�l�擪�ʒu�A�h���X�i�[�p�ϐ�
    Dim category_count As Long      ' �񓚐��i�[�p�ϐ�
    
    Dim work_count As Long          ' �񓚐��J�E���g�p�ϐ�
    
    Dim target_address As String    ' �������R�[�h�擪�A�h���X�i�[�p�ϐ�
    Dim ct_count As Long            ' �����ʒu�i�[�p�ϐ�
    Dim lighting_count As Long      ' �������J�E���g�p�ϐ�
    
    Dim input_type As Long          ' ���~�b�g�̖��񓚂ɑ΂�����͌`�Ԕ���p�ϐ�
                                    ' [1] 0 input [2] "" input
    
    ' �ő�v�f�����擾
    process_max = UBound(q_data, 1)
    
    ' �ő���H�񐔕��������s��
    For process_count = 1 To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text _
        '& "�@���~�b�g�}���`���H������(" & Format(process_count - START_SUTATUSBER) & "/" & _
        'Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
        
        ' �t�H�[�}�b�g��LM�̎�
        If Mid(q_data(process_count).q_format, 1, 1) = "L" Then
            
            ' ���̓f�[�^�S�Ă𔻒肷��
            For indata_count = START_ROW_INDATA To indata_maxrow
        
                ' �J�E���g�����񓚐���������
                lighting_count = 0
        
                ' ��ƃG���A�̐擪�A�h���X���擾
                target_address = ws_process.Cells(indata_count, q_data(process_count).data_column).Address
        
                ' �Ώ۔͈͓��̉񓚐����擾�A���[�v�J�E���g���������v�f���̐ݖ�݂̂ɏ������s��
                If WorksheetFunction.CountIf(wsp_indata.Range(target_address).Resize(, q_data(process_count).ct_count), ">0") > _
                q_data(process_count).ct_loop Then
                
                    ' �L�����e�ɂ�蔻��
                    Select Case q_data(process_count).q_format
                    
                        ' �f�[�^���N���A����
                        Case "LC"
                        
                            ' �f�[�^��������
                            wsp_indata.Range(target_address).Resize(, q_data(process_count).ct_count).ClearContents
        
                        ' ���ԗD��ŃJ�e�S���[���c��
                        Case "LA"
                        
                            ' ��ԍ���D��ŏ������s��
                            For ct_count = q_data(process_count).ct_count To 1 Step -1
                        
                                ' ���[�v�J�E���g���܂ł̉񓚐����J�E���g
                                If lighting_count < q_data(process_count).ct_loop Then
                        
                                    ' �񓚓��e���m�F����
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' �񓚐��𑝉�������
                                        lighting_count = lighting_count + 1
                        
                                    End If
                                    
                                ' ���[�v�J�E���g�ȍ~�̃Z��
                                Else
                        
                                    ' �f�[�^��������
                                    wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                        
                                End If
                        
                            Next ct_count
        
                        ' ��ԗD��ŃJ�e�S���[���c��
                        Case "L", "LM"
                            
                            ' ��ԍ���D��ŏ������s��
                            For ct_count = 1 To q_data(process_count).ct_count
                        
                                ' ���[�v�J�E���g���܂ł̉񓚐����J�E���g
                                If lighting_count < q_data(process_count).ct_loop Then
                        
                                    ' �񓚓��e���m�F����
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' �񓚐��𑝉�������
                                        lighting_count = lighting_count + 1
                        
                                    End If
                                    
                                ' ���[�v�J�E���g�ȍ~�̃Z��
                                Else
                        
                                    ' �f�[�^��������
                                    wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                        
                                End If
                        
                            Next ct_count
                        
                        ' �w�����Ȃ��ꍇ�͉����s��Ȃ�
                        Case Else
        
                    End Select
                
                End If
        
            Next indata_count
        
        End If

    Next process_count
    
End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.05.02  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.02�@'
' ���H�ő}��������̃w�b�_�����쐬����֐�                                                       '
' �����P WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����Q String�^ input_code                                                                       '
' �����R String�^ hedder_address                                                                   '
' �߂�l boolean�^ �����̗L��                                                                      '
'--------------------------------------------------------------------------------------------------'
Private Function Hedder_Create(ByVal wsp_indata As Worksheet, ByVal input_code As String, _
ByVal hedder_address As String) As Boolean
        
        Dim color_column As Long    ' ���F�p�J�����i�[�p�ϐ�
        Dim qcode_row As Long       ' QCODE��ԍ��i�[�p�ϐ�
        Dim ma_ct As Long           ' MA�J�e�S���[���i�[�p�ϐ�
        Dim loop_count As Long      ' ������J�E���g�p�ϐ�
        
        ' QCODE������
        qcode_row = Qcode_Match(input_code)
        
        ' MA��LM�̎�
        If Mid(q_data(qcode_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode_row).q_format, 1, 1) = "L" Then
            ' CT�����i�[
            ma_ct = q_data(qcode_row).ct_count
        Else
            ' MA��LM�ȊO��1��ݒ�
            ma_ct = 1
        End If
        
    ' �w��񕪏������s��
    For loop_count = 1 To ma_ct
        
        ' �R�����g����̓f�[�^�w�b�_�ɑ���A����������
        wsp_indata.Range(hedder_address).Value = input_code
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        
        ' �w�b�_�A�h���X��ύX
        hedder_address = wsp_indata.Range(hedder_address).Offset(1).Resize(3).Address
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        
        ' �S�̂Ɍr��������
        hedder_address = wsp_indata.Range(hedder_address).Offset(3).Resize(2).Address
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlDash
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).Weight = xlHairline
        
        ' �͈͂�ύX
        hedder_address = wsp_indata.Range(hedder_address).Offset(-4).Resize(6).Address
        
        ' �t�H�[�}�b�g�ɂ���ď������s��
        Select Case Mid(q_data(qcode_row).q_format, 1, 1)
            Case "S"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "SA"
            Case "M"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(1).Resize(1) = loop_count
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "MA"
            Case "L"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(1).Resize(1) = loop_count
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "LM"
            Case "R"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "RA"
            Case "H"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("�ݒ���").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "HC"
           ' �b�蒅�F�̂��߁A�w�肷��ꍇ�͎d�l�����߂Ē��F���邱��
           Case Else
                wsp_indata.Range(hedder_address).Interior.Color = RGB(255, 192, 0)
        End Select
        
        ' �b�s�ԍ��ƃt�H�[�}�b�g���Z���^�����O
        wsp_indata.Range(hedder_address).Offset(1).Resize(1).HorizontalAlignment = xlCenter
        wsp_indata.Range(hedder_address).Offset(3).Resize(1).HorizontalAlignment = xlCenter
        
        ' �����ʒu��1�s���炷
        hedder_address = wsp_indata.Range(hedder_address).Offset(0, 1).Resize(1).Address
         
    Next loop_count

End Function

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.05.09  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.15�@'
' QCODE�̓��̓f�[�^���w���ʂ�ɋL������Ă��邩�𔻒肷��v���V�[�W��                              '
' �����P WorkSheet�^ wsp_indata ���̓f�[�^�V�[�g                                                   '
' �����Q Variant�^   and_data ���H�w���s�i�[�p�z��ϐ�(�L�����S��)                               '
' �����R Variant�^   and_column ���H�w���s�ԍ��i�[�p�z��ϐ�(Long�^)                               '
' �����S Long�^      indata_maxrow ���̓f�[�^�ŏI��ԍ��i�[�p�ϐ�                                  '
' �����T Long�^      qcode2_row �o�͐��ԍ��i�[�p�ϐ�                                             '
' �߂�l boolean�^   salectflg_decision �����̗L��                                                 '
'--------------------------------------------------------------------------------------------------'
Private Sub salectflg_decision(ByVal wsp_indata As Worksheet, ByVal and_data As Variant, _
ByVal and_column As Variant, ByVal indata_maxrow As Long, ByVal qcode2_row As Long)

    Dim pindata_count As Long       ' ���̓f�[�^�J�E���g�p�ϐ�
    Dim loop_count As Long          ' ���[�v�J�E���g�p�ϐ�
    
    Dim decision_flg1 As Boolean    ' ��������p�t���O�P
    Dim decision_flg2 As Boolean    ' ��������p�t���O�Q
    Dim decision_flg3 As Boolean    ' ��������p�t���O�R
    
    Dim decision_type As Long       ' AND�AOR�����i�[�p�ϐ�
    Dim qcode1_row As Long          ' ���̓f�[�^�ʒu�i�[�p�ϐ�
    Dim min_number As Double        ' �Z���N�g�����ŏ��l�i�[�p�ϐ�
    Dim max_number As Double        ' �Z���N�g�����ő�l�i�[�p�ϐ�
    
    Dim ma_address As String        ' �����񓚃A�h���X�i�[�p�ϐ�
    
    ' �S�Ă̓��̓f�[�^�ɏ������s��
    For pindata_count = START_ROW_INDATA To indata_maxrow
    
        'MsgBox wsp_indata.Cells(pindata_count, 1).Value
    
        ' ������������p�t���O��������
        decision_flg1 = False
        decision_flg2 = True
        decision_flg3 = False
        
        ' �z��Ɋi�[�������H�w�����̏������s��
        For loop_count = 1 To UBound(and_column)
                
            ' �����񓚂̔���������擾�i���̏����j
            If UBound(and_data, 2) > 11 + ((loop_count - 1) * 6) + 1 And decision_flg2 = True Then
            ' ��b���11�s�A�ǉ���񖈂�6�s�A���̏����͂���1�s��

                ' OR
                If and_data(1, and_column(loop_count + 1)) = "or (��������)" Then
                    decision_type = 1
                ' AND
                ElseIf and_data(1, and_column(loop_count + 1)) = "and (����)" Then
                    decision_type = 2
                ' ���̑��i���󎟂̏����������Ƃ��̂݁j
                Else
                    decision_type = 3
                End If
            
            ' ���̏�����������
            Else
                decision_type = 3
            End If
            
            ' QCODE�̗�ԍ����擾
            qcode1_row = Qcode_Match(and_data(1, and_column(loop_count) + 1))
            min_number = and_data(1, and_column(loop_count) + 2)
            max_number = and_data(1, and_column(loop_count) + 4)
            
            ' �w��t�H�[�}�b�g�ɍ��킹�ď������s��
            Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                ' �P��񓚂̏ꍇ
                Case "S", "R", "H"
                    
                    ' �w��̃Z����񂪎w��͈̔͂̋L���ł��鎞
                    If wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column) >= min_number And _
                    wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column) <= max_number Then
                    
                        ' �Z���N�g�t���O��L��
                        decision_flg1 = True    ' �Z���N�g�L���t���O
                        
                        ' ���̐ڑ����ɍ��킹�ăt���O��ύX����A�p�������t���O��L��
                        Select Case decision_type
                        
                            ' OR����
                            Case 1
                                decision_flg2 = True    ' �p�������t���O
                                decision_flg3 = True    ' OR�����J�n�t���O
                            ' AND����
                            Case 2
                                decision_flg2 = True    ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                            ' ����
                            Case Else
                                decision_flg2 = False   ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                        End Select
                    
                    ' �͈͊O�ł������ꍇ
                    Else
                    
                        ' ��O�̏�����OR�����ł͂Ȃ��A�����̏�����OR�ł͖�����
                        If decision_flg3 = False And decision_type <> 1 Then
                    
                            ' �Z���N�g�����𖞂����Ȃ����߃t���O��S��FALSE�ɂ��I��
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = False   ' �p�������t���O
                            decision_flg3 = False   ' OR�����J�n�t���O
                            Exit For
                        
                        ' ��O�̏�����OR�ȊO�ŁA�����̏�����OR�̎�
                        ElseIf decision_flg3 = False And decision_type = 1 Then
                        
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = True    ' �p�������t���O
                            decision_flg3 = True    ' OR�����J�n�t���O
                        
                        ' ��O�̏�����OR����
                        ElseIf decision_flg3 = True Then
                        
                            ' �Z���N�g�L���񓚃t���O���I���̎��iOR�����̏����𖞂����Ă���ꍇ
                            If decision_flg1 = True Then
                            
                                ' ���̏�����OR�̎�
                                If decision_type = 1 Then
                            
                                    decision_flg2 = True   ' �p�������t���O
                                    decision_flg3 = True   ' OR�����J�n�t���O
                            
                                ' ���̏�����AND�̎�
                                ElseIf decision_type = 2 Then
                            
                                    decision_flg2 = True   ' �p�������t���O
                                    decision_flg3 = False  ' OR�����J�n�t���O
                                
                                ' ����ȊO�̎�
                                Else
                                
                                    decision_flg2 = False  ' �p�������t���O
                                    decision_flg3 = False  ' OR�����J�n�t���O
                                    Exit For
                                
                                End If
                            
                            ' ���̏�����OR�̎�
                            ElseIf decision_type = 1 Then
                                
                                decision_flg2 = True   ' �p�������t���O
                                decision_flg3 = True   ' OR�����J�n�t���O
                                
                            Else
                                
                                decision_flg1 = False   ' �Z���N�g�L���t���O
                                decision_flg2 = False   ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                                Exit For
                                
                            End If
                            
                        ' ���̏�����OR�����̎�
                        ElseIf decision_type = 1 Then
                            
                            decision_flg2 = True   ' �p�������t���O
                            decision_flg3 = True   ' OR�����J�n�t���O
                                
                        ' ����ȊO�̎�
                        Else
                                
                            ' �Z���N�g�����𖞂����Ȃ����߃t���O��S��FALSE�ɂ��I��
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = False   ' �p�������t���O
                            decision_flg3 = False   ' OR�����J�n�t���O
                                
                        End If
                        
                    End If
                
                ' �����񓚂̏ꍇ
                Case "M", "L"
                
                    ' �����ʒu�̐擪�A�h���X���i�[
                    ma_address = wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column).Address
                    
                    ' 0�J�e�S���[���̂ݎQ�ƈʒu�𒲐�����
                    'If q_data(qcode1_row).ct_0flg = True Then
                    '
                    '    min_number = min_number + 1
                    '    max_number = max_number + 1
                    '
                    'End If
                    
                    ' �w��̃Z����񂪎w��͈̔͂̋L���ł��鎞
                    If WorksheetFunction.Sum(wsp_indata.Range(ma_address).Offset(0, min_number - 1) _
                    .Resize(, max_number - min_number + 1)) <> 0 Then
                    
                        ' �Z���N�g�t���O��L��
                        decision_flg1 = True    ' �Z���N�g�L���t���O
                        
                        ' ���̐ڑ����ɍ��킹�ăt���O��ύX����A�p�������t���O��L��
                        Select Case decision_type
                        
                            ' OR����
                            Case 1
                                decision_flg2 = True    ' �p�������t���O
                                decision_flg3 = True    ' OR�����J�n�t���O
                            ' AND����
                            Case 2
                                decision_flg2 = True    ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                            ' ����
                            Case Else
                                decision_flg2 = False   ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                        End Select
                    
                    ' �͈͊O�ł������ꍇ
                    Else
                    
                        ' ��O�̏�����OR�����ł͂Ȃ��A�����̏�����OR�ł͖�����
                        If decision_flg3 = False And decision_type <> 1 Then
                    
                            ' �Z���N�g�����𖞂����Ȃ����߃t���O��S��FALSE�ɂ��I��
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = False   ' �p�������t���O
                            decision_flg3 = False   ' OR�����J�n�t���O
                            Exit For
                        
                        ' ��O�̏�����OR�ȊO�ŁA�����̏�����OR�̎�
                        ElseIf decision_flg3 = False And decision_type = 1 Then
                        
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = True    ' �p�������t���O
                            decision_flg3 = True    ' OR�����J�n�t���O
                        
                        ' ��O�̏�����OR����
                        ElseIf decision_flg3 = True Then
                        
                            ' �Z���N�g�L���񓚃t���O���I���̎��iOR�����̏����𖞂����Ă���ꍇ
                            If decision_flg1 = True Then
                            
                                ' ���̏�����OR�̎�
                                If decision_type = 1 Then
                            
                                    decision_flg2 = True   ' �p�������t���O
                                    decision_flg3 = True   ' OR�����J�n�t���O
                            
                                ' ���̏�����AND�̎�
                                ElseIf decision_type = 2 Then
                            
                                    decision_flg2 = True   ' �p�������t���O
                                    decision_flg3 = False  ' OR�����J�n�t���O
                                
                                ' ����ȊO�̎�
                                Else
                                
                                    decision_flg2 = False  ' �p�������t���O
                                    decision_flg3 = False  ' OR�����J�n�t���O
                                    Exit For
                                
                                End If
                            
                            ' ���̏�����OR�̎�
                            ElseIf decision_type = 1 Then
                                
                                decision_flg2 = True   ' �p�������t���O
                                decision_flg3 = True   ' OR�����J�n�t���O
                                
                            Else
                                
                                decision_flg1 = False   ' �Z���N�g�L���t���O
                                decision_flg2 = False   ' �p�������t���O
                                decision_flg3 = False   ' OR�����J�n�t���O
                                Exit For
                                
                            End If
                            
                        ' ���̏�����OR�����̎�
                        ElseIf decision_type = 1 Then
                            
                            decision_flg2 = True   ' �p�������t���O
                            decision_flg3 = True   ' OR�����J�n�t���O
                                
                        ' ����ȊO�̎�
                        Else
                                
                            ' �Z���N�g�����𖞂����Ȃ����߃t���O��S��FALSE�ɂ��I��
                            decision_flg1 = False   ' �Z���N�g�L���t���O
                            decision_flg2 = False   ' �p�������t���O
                            decision_flg3 = False   ' OR�����J�n�t���O
                                
                        End If
                        
                    End If
        
                Case Else
        
            End Select
            
            ' �p���t���O�������̏ꍇ
            If decision_flg2 = False Then
                Exit For
            End If
            
        Next loop_count
        
        ' �Z���N�g�t���O���L���̏ꍇ
        If decision_flg1 = True Then
        
            ' �Z���N�g�t���O��L���ɂ���
            wsp_indata.Cells(pindata_count, q_data(qcode2_row).data_column).Value = 1
        
        End If
        
    Next pindata_count

    'Debug.Print and_data(1, and_column(1))

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.05.15  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.19�@'
' �w�肳�ꂽ�񓚂�����ꍇ�f�[�^���N���A����                                                       '
' �����P WorkSheet�^ �t�Z�b�g�����w���V�[�g                                                        '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub data_clear1(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long             ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long              ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long              ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim input_word As String            ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean          ' ��������p�t���O
    
    Dim clear_min As Long               ' �폜�l�i�[�p�ϐ��i�ŏ��l�j
    Dim clear_max As Long               ' �폜�l�i�[�p�ϐ��i�ő�l�j
    
    Dim indata_count As Long            ' ���̓f�[�^�����ʒu�i�[�p�ϐ�

    Dim ma_count As Long                ' MA�񓚓��e�m�F�p�J�E���g�ϐ�
    
    ' 20200331 �ǉ�
    Dim amplification_count As Long     ' �������J�E���g�p�ϐ�
    
    Dim qcode_count As Long             ' �ݒ��ʏ��J�E���g�p�ϐ�
    Dim qcode_max As Long               ' �ݒ��ʏ��ő吔�i�[�p�ϐ�
    Dim select_qcode(2, 3) As Long      ' �Z���N�g����QCODE�i�[�z��
    
    Dim processing_count As Long        ' �Z���N�g�������J�E���g�p�ϐ�
    Dim processing_data As Long         ' �Z���N�g�������i�[�p�ϐ�
    Dim processing_flg As Boolean       ' �Z���N�g��������p�ϐ�
    Dim processing_address As String    ' �Z���N�g�����A�h���X�i�[�p�ϐ�
    
    Dim target_data As String           ' �f�[�^�N���A�A�h���X�i�[�p�ϐ�
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�f�[�^�N���A�@������(" & Format(process_count - START_SUTATUSBER) & _
        '"/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �X�L�b�v�t���O���L���łȂ����AQCODE���L������Ă��鎞
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA3) <> "" Then
        
            ' �Q�Ɛݖ��ԍ����擾
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2).Value)
            clear_min = ws_process.Cells(process_count, QCODE1_DATA1).Value
            clear_max = ws_process.Cells(process_count, QCODE1_DATA3).Value
            
            ' ���̓f�[�^�S�Ăɏ������s��
            For indata_count = START_ROW_INDATA To indata_maxrow
            
                ' ��������p�t���O������
                process_flg = False
                
                ' �Q�ƃG���A�̃A�h���X���擾
                target_data = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
            
                ' �Q�Ɛݖ�̃t�H�[�}�b�g�ɂ�菈����ύX����
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' SA
                    Case "S", "R", "H"
                    
                        ' �w��͈͂̋L��������ꍇ
                        If wsp_indata.Range(target_data) >= clear_min And _
                        wsp_indata.Range(target_data) <= clear_max Then
                
                            ' ��������t���O��L���ɂ���
                            process_flg = True
                
                        End If
                    
                    ' MA LM
                    Case "M", "L"
                        
                        ' ct_0flg��ON�̎��͍��W��1�ύX����
                        'If q_data(qcode1_row).ct_0flg = True Then
                        '    clear_min = clear_min + 1
                        '    clear_max = clear_max + 1
                        'End If
                        
                        ' �w��͈͂̋L��������ꍇ
                        If Application.WorksheetFunction.Sum(wsp_indata.Range(target_data). _
                        Offset(, clear_min - 1).Resize(, clear_max)) <> 0 Then
                            
                            ' ��������p�t���O��L���ɂ���
                            process_flg = True
                        
                        End If
                        
                    Case Else
            
                End Select
            
                ' ��������p�t���O���L���̏ꍇ
                If process_flg = True Then
                
                    ' �f�[�^�N���A�ݖ�̃t�B�[�}�b�g�ɂ�菈�����킯��
                    Select Case Mid(q_data(qcode2_row).q_format, 1, 1)
                        
                        ' SA
                        Case "S", "R", "H", "F", "O"
                        
                            ' �f�[�^�N���A�ݖ���N���A
                            wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Clear
                
                        ' MA LM
                        Case "M", "L"
                
                            wsp_indata.Cells(wsp_indata, q_data(qcode2_row).data_column). _
                            Resize(, q_data(qcode2_row).ct_count).Clear
                
                        Case Else
                
                    End Select
                
                End If
            
            Next indata_count
        
        End If
    
    Next process_count

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.05.15  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.06.26�@'
' �ݒu��ʂ̃Z���N�g�����ɍ��킹�ē��̓f�[�^���N���A����                                           '
' �����P WorkSheet�^ �t�Z�b�g�����w���V�[�g                                                        '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub data_clear2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long             ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long              ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long              ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim input_word As String            ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean          ' ��������p�t���O
    
    Dim indata_count As Long            ' ���̓f�[�^�����ʒu�i�[�p�ϐ�

    Dim ma_count As Long                ' MA�񓚓��e�m�F�p�J�E���g�ϐ�
    
    Dim qcode_count As Long             ' �ݒ��ʏ��J�E���g�p�ϐ�
    Dim qcode_max As Long               ' �ݒ��ʏ��ő吔�i�[�p�ϐ�
    Dim select_qcode(2, 3) As Long      ' �Z���N�g����QCODE�i�[�z��
    
    Dim processing_count As Long        ' �Z���N�g�������J�E���g�p�ϐ�
    Dim processing_data As Long         ' �Z���N�g�������i�[�p�ϐ�
    Dim processing_flg As Boolean       ' �Z���N�g��������p�ϐ�
    Dim processing_address As String    ' �Z���N�g�����A�h���X�i�[�p�ϐ�
    
    Dim target_data As String           ' �f�[�^�N���A�A�h���X�i�[�p�ϐ�
    
    ' QCODE����S�Ċm�F����
    For qcode_count = 1 To UBound(q_data, 1)
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text _
        '& "�@�f�[�^�N���A�A������(" & Format(qcode_count) & "/" & Format(UBound(q_data)) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �Z���N�g�����@�����݂��鎞
        If q_data(qcode_count).sel_code1 <> "" Then
    
            ' �Z���N�g�����ɍ��킹�N���A�������s��
            Call select_clear(Qcode_Match(q_data(qcode_count).sel_code1), q_data(qcode_count).sel_value1, qcode_count, wsp_indata, indata_maxrow)
            
            ' �Z���N�g�����A�����݂��鎞
            If q_data(qcode_count).sel_code2 <> "" Then
            
                ' �Z���N�g�����ɍ��킹�N���A�������s��
                Call select_clear(Qcode_Match(q_data(qcode_count).sel_code2), q_data(qcode_count).sel_value2, qcode_count, wsp_indata, indata_maxrow)
                
                ' �Z���N�g�����B�����݂��鎞
                If q_data(qcode_count).sel_code3 <> "" Then
                
                    ' �Z���N�g�����ɍ��킹�N���A�������s��
                    Call select_clear(Qcode_Match(q_data(qcode_count).sel_code3), q_data(qcode_count).sel_value3, qcode_count, wsp_indata, indata_maxrow)
                
                End If
            
            End If
    
        End If
            
    Next qcode_count

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.05.15  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.06.26�@'
' ���~�b�g�}���`�̉񓚏󋵏C�������@�@  �@�@�@�@�@�@�@�@�@                                         '
' �����P WorkSheet�^ �t�Z�b�g�����w���V�[�g                                                        '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Limitmulti_2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long             ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long              ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long              ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim input_word As String            ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean          ' ��������p�t���O
    
    Dim indata_count As Long            ' ���̓f�[�^�����ʒu�i�[�p�ϐ�

    Dim ma_count As Long                ' MA�񓚓��e�m�F�p�J�E���g�ϐ�
    
    Dim qcode_count As Long             ' �ݒ��ʏ��J�E���g�p�ϐ�
    Dim qcode_max As Long               ' �ݒ��ʏ��ő吔�i�[�p�ϐ�
    Dim select_qcode(2, 3) As Long      ' �Z���N�g����QCODE�i�[�z��
    
    Dim processing_count As Long        ' �Z���N�g�������J�E���g�p�ϐ�
    Dim processing_data As Long         ' �Z���N�g�������i�[�p�ϐ�
    Dim processing_flg As Boolean       ' �Z���N�g��������p�ϐ�
    Dim processing_address As String    ' �Z���N�g�����A�h���X�i�[�p�ϐ�
    
    Dim target_data As String           ' �f�[�^�N���A�A�h���X�i�[�p�ϐ�
    
'    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
'    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
'    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    
'    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
'    Dim qcode2_dataflg As Boolean   ' �t�Z�b�g�Ώېݖ�(�e)����t���O
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu

    Dim process_case As Long        ' �������e�i�[�p�ϐ�
    Dim start_address As String     ' �k�l�擪�ʒu�A�h���X�i�[�p�ϐ�
    Dim category_count As Long      ' �񓚐��i�[�p�ϐ�
    
    Dim work_count As Long          ' �񓚐��J�E���g�p�ϐ�
    
    Dim target_address As String    ' �������R�[�h�擪�A�h���X�i�[�p�ϐ�
    Dim ct_count As Long            ' �����ʒu�i�[�p�ϐ�
    Dim lighting_count As Long      ' �������J�E���g�p�ϐ�
    
    Dim search_area As Range        ' 1.0 or 1."" ����T�[�`�p�ϐ�
    Dim search_address As Range     ' �T�[�`���i�[�p�ϐ�
    Dim search_flg As Boolean       ' �T�[�`����t���O�itrue = 0�A���Afalse = 0�i�V�j
    
    ' QCODE����S�Ċm�F����
    For qcode_count = 1 To UBound(q_data, 1)
    
        ' �t�H�[�}�b�g��LM�̎�
        If Mid(q_data(qcode_count).q_format, 1, 1) = "L" Then
        
            ' �w��̃��~�b�g�}���`��1.0 or 1.""���𔻒�
            Set search_area = Range(ws_process.Cells(START_ROW_INDATA, q_data(qcode_count).data_column).Address, _
                ws_process.Cells(indata_maxrow, q_data(qcode_count).data_column + q_data(qcode_count).ct_count - 1).Address)
            Set search_address = search_area.Find(0, LookIn:=xlValues, lookat:=xlWhole)
            
            ' ����t���O�؂�ւ�
            If Not search_address Is Nothing Then
                search_flg = True
            Else
                search_flg = False
            End If
            
            ' ���̓f�[�^�S�Ă𔻒肷��
            For indata_count = START_ROW_INDATA To indata_maxrow
        
                ' �J�E���g�����񓚐���������
                lighting_count = 0
        
                ' ��ƃG���A�̐擪�A�h���X���擾
                target_address = ws_process.Cells(indata_count, q_data(qcode_count).data_column).Address
        
                ' �Ώ۔͈͓��̉񓚐����擾�A���[�v�J�E���g���������v�f���̐ݖ�݂̂ɏ������s��
                If WorksheetFunction.CountIf(wsp_indata.Range(target_address).Resize(, q_data(qcode_count).ct_count), ">0") > _
                q_data(qcode_count).ct_loop Then
                
                    ' �L�����e�ɂ�蔻��
                    Select Case q_data(qcode_count).q_format
                    
                        ' �f�[�^���N���A����
                        Case "LC"
                        
                            ' �f�[�^��������
                            wsp_indata.Range(target_address).Resize(, q_data(qcode_count).ct_count).ClearContents
        
                        ' ���ԗD��ŃJ�e�S���[���c��
                        Case "LA"
                        
                            ' ��ԍ���D��ŏ������s��
                            For ct_count = q_data(qcode_count).ct_count To 1 Step -1
                        
                                ' ���[�v�J�E���g���܂ł̉񓚐����J�E���g
                                If lighting_count < q_data(qcode_count).ct_loop Then
                        
                                    ' �񓚓��e���m�F����
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' �񓚐��𑝉�������
                                        lighting_count = lighting_count + 1
                                    
                                    ' 1.0�`���̏ꍇ�A0�E��
                                    ElseIf search_flg = True Then
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1) = 0
                                    End If
                                    
                                ' ���[�v�J�E���g�ȍ~�̃Z��
                                Else
                                    If search_flg = True Then
                                        ' �f�[�^�������� 1
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    Else
                                        ' �f�[�^�������� ""
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                                    End If

                        
                                End If
                        
                            Next ct_count
        
                        ' ��ԗD��ŃJ�e�S���[���c��
                        Case "L", "LM"
                            
                            ' ��ԍ���D��ŏ������s��
                            For ct_count = 1 To q_data(qcode_count).ct_count
                        
                                ' ���[�v�J�E���g���܂ł̉񓚐����J�E���g
                                If lighting_count < q_data(qcode_count).ct_loop Then
                        
                                    ' �񓚓��e���m�F����
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' �񓚐��𑝉�������
                                        lighting_count = lighting_count + 1
                                    
                                    ' 1.0�`���̏ꍇ�A0�E��
                                    ElseIf search_flg = True Then
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    End If
                                    
                                ' ���[�v�J�E���g�ȍ~�̃Z��
                                Else
                                    '
                                    If search_flg = True Then
                                        ' �f�[�^�������� 1
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    
                                    Else
                                        ' �f�[�^�������� ""
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                                    End If
                                    
                                End If
                        
                            Next ct_count
                        
                        ' �w�����Ȃ��ꍇ�͉����s��Ȃ�
                        Case Else
        
                    End Select
                
                End If
        
            Next indata_count
        
        End If
            
    Next qcode_count

End Sub

'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2017.05.08  '
' �Z���N�g�t���O���H�p�v���V�[�W��                                                                 '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim qcode3_row As Long          ' �G���g���[�G���A�I�[�i�[�p�ϐ�
    
    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu
    
    Dim str_ct1 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct2 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct3 As Variant          ' �J�e�S���[��r�p�z��i��+���j
    
    Dim ct_count As Long            ' �z����e�J�E���g�p�ϐ�
    
'    Dim ct1_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct2_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct3_count As Long           ' ��+���񓚐��i�[�p�ϐ�
    
    Dim work1_flg As Boolean        ' �������i�[�p�t���O
    Dim work2_flg As Boolean        ' �������i�[�p�t���O
    
    Dim target_address As String    ' QCODE1�A�h���X�i�[�p�ϐ�
    
    Dim processing_flg As Boolean   ' Function�߂�l�i�[�p�ϐ�
    
    Dim wb_calculation As Workbook  ' ���Ԍv�Z�p�u�b�N���i�[�p�ϐ�
    Dim ws_calculation As Worksheet ' ���Ԍv�Z�p�V�[�g���i�[�p�ϐ�
    
    Dim start_num As Double         ' �J�n�ԍ��i�[�p�ϐ�
    Dim end_num As Double           ' �I�[�ԍ��i�[�p�ϐ�
    Dim connect_data As String      ' �ڑ����i�[�p�ϐ�
    Dim target_range As Long        ' �Z���N�g�͈͎擾�p�ϐ�
    Dim answer_num As Long          ' �񓚐��i�[�p�ϐ�
    Dim match_data As Long          ' �쐬QCODE�s�ԍ��i�[�p�ϐ�
    
    Dim match_column As Long
    
' �e�w���P�ꖈ�ɏ����𖞂����Ă��邩�𔻒肵�ĕʃu�b�N�ɓf���o��----------------------------------------------
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �v�Z�p�u�b�N���쐬
    Set wb_calculation = Workbooks.Add
    Set ws_calculation = wb_calculation.Worksheets(1)
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�Z���N�g�t���O���H������(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �X�L�b�v�t���O�����͂���Ă��Ȃ���
        If Len(ws_process.Cells(process_count, SKIP_FLG).Value) = 0 Then
        
            ' �Q�Ɛݖ��QCODE�̗�ԍ����擾
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4).Value)
            qcode3_row = Qcode_Match("*���H��")
            
            ' ���ɏo�̓G���A���p�ӂ���Ă��鎞�i�����H�������ɏo�̓G���A���ݒ肳��Ă��鎞
            If q_data(qcode2_row).data_column <> 0 And _
            q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
        
            ' �܂��G���A���p�ӂ���Ă��Ȃ���
            Else
        
                ' �w�b�_���쐬����
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' �V�����ݒ肵���G���A��q_data�ɃJ�����Ƃ��Đݒ肷��
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' �ŏ��l�A�ő�l����̓f�[�^�̃w�b�_�Ɋi�[
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' �v�Z�V�[�g�ɓ���QCODE��������
            ws_calculation.Cells(START_ROW - 3, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1)
            
            ' �v�Z�V�[�g�ɏo��QCODE��������
            ws_calculation.Cells(START_ROW - 2, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA4)
            
            ' �v�Z�V�[�g�ɏo�͐�QCODE��ԍ���������
            ws_calculation.Cells(START_ROW - 1, process_count - QCODE_P_COLUMN).Value = qcode2_row
            
            ' �v�Z�V�[�g�ɏW�v������������
            ws_calculation.Cells(START_ROW, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA5)
            
            ' �J�n�ԍ��ƏI�[�ԍ���ϐ��ɕێ�������
            start_num = ws_process.Cells(process_count, QCODE1_DATA1).Value
            end_num = ws_process.Cells(process_count, QCODE1_DATA3).Value
            
            ' ���̓f�[�^�S�Ăɏ������s��
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' �Ώېݖ�̃t�H�[�}�b�g�ɂ�蔻�������ύX
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' �P��񓚂̏ꍇ
                    Case "S", "R", "H"
                
                        ' �񓚂��w�肳�ꂽ�͈͂Ɋ܂܂�Ă��邩����
                        If wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value >= start_num And _
                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <= end_num Then
                            
                            ' �t���O�𗧂Ă�
                            ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            
                        End If
                    
                    ' �����񓚂̏ꍇ
                    Case "L", "M"
                    
                        ' �擪�A�h���X���擾
                        target_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                        
                        ' 0CT�t���O�𔻒�i�ʏ�J�e�S���[�j
                        'If q_data(qcode1_row).ct_0flg = False Then
                        
                            ' �Ώۂ͈̔͂̉񓚂��m�F����
                            If WorksheetFunction.Sum(wsp_indata.Range(target_address) _
                            .Offset(0, start_num - 1).Resize(, end_num + 1 - start_num)) <> 0 Then
                        
                                ' �t���O�𗧂Ă�
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                            End If
                        
                        ' 0CT�t���O�𔻒�i�O�J�e�S���[�L��j
                        'Else
                        
                            ' �Ώۂ͈̔͂̉񓚂��m�F����
                        '    If WorksheetFunction.sum(wsp_indata.Range(target_address) _
                        '    .Offset(0, start_num).Resize(, (end_num + 1) - start_num)) <> 0 Then
                        
                                ' �t���O�𗧂Ă�
                        '        ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                        '    End If
                        
                        'End If

                    Case Else
                    
                    End Select
                    
            Next indata_count
        
        End If
    
    Next process_count
    
' �ʃu�b�N�ɓf���o����������𐮌`���Z���N�g�t���O���쐬����---------------------------------------------
    
    ' �v�Z�V�[�g���A�N�e�B�u�ɕύX
    ws_calculation.Activate
    
    ' �ő���H�񐔕��������s��
    For process_count = (START_ROW - QCODE_P_COLUMN) To (process_max - QCODE_P_COLUMN)
    
        ' ��ʂւ̕\�����I���ɂ���
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"�@�Z���N�g�t���O���H������(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' ��ʂւ̕\�����I�t�ɂ���
        'Application.ScreenUpdating = False
    
        ' �O��̎w���Əo��QCODE���قȂ鎞
        If ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count + 1).Value And _
        ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count - 1).Value Then
        
            ' �v�Z�V�[�g���A�N�e�B�u�ɕύX
            ws_calculation.Activate
        
            ' �ʏ폈��
            target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            
            ' �Ώۂ͈̔͂��R�s�[���ē��̓f�[�^�ɓ\��t����
            ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
            Selection.Copy
            
            ' ���̓f�[�^���A�N�e�B�u�ɕύX
            wsp_indata.Activate
            
            ' �������݈ʒu�̃A�h���X���擾
            target_address = wsp_indata.Cells(START_ROW_INDATA, q_data(ws_calculation.Cells(QCODE_P_ROW, process_count).Value).data_column).Address
            
            ' �R�s�[�����f�[�^��\��t��
            wsp_indata.Range(target_address).PasteSpecial (xlPasteValues)
            
        ' �O�̎w���Əo��QCODE���قȂ鎞�i�n�_�j
        ElseIf ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count - 1).Value Then
        
            ' �n�_�̃A�h���X�A�w���`�Ԃ��擾
            target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            connect_data = ws_calculation.Cells(QCODE_P_CA, process_count).Value
            target_range = 1
            
            ' �w�b�_�����R�s�[����
            ws_calculation.Cells(QCODE_P_COLUMN, 1).Value = ws_calculation.Cells(QCODE_P_COLUMN, process_count).Value
            ws_calculation.Cells(QCODE_P_ROW, 1).Value = ws_calculation.Cells(QCODE_P_ROW, process_count).Value
        
        ' ��̎w���Əo��QCODE���قȂ鎞�i�I�_�j
        ElseIf ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count + 1).Value Then
        
            ' �J�E���^�[�𑝉�����
            target_range = target_range + 1
        
            ' ���̓f�[�^�S�Ăɏ������s��
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' �w��͈͂̉񓚐����擾
                answer_num = WorksheetFunction.Sum(ws_calculation.Range(target_address) _
                .Offset(indata_count - START_ROW - 1, 0).Resize(, target_range))
            
                ' �����ɍ��킹�ď������s��
                Select Case connect_data
            
                    ' OR
                    Case "or (��������)"
                    
                        ' �Ώ۔͈͂Ɉ�ł��񓚂�������
                        If answer_num <> 0 Then
                            
                            ' �t���O��L���ɂ���
                            ws_calculation.Cells(indata_count, 1).Value = 1
                        
                        Else
                        
                            ' �L���o�Ȃ��ꍇ�̓N���A���s��
                            ws_calculation.Cells(indata_count, 1).Value = ""
                        
                        End If
                        
                    ' AND
                    Case "and (����)"
                    
                        ' �Ώ۔͈͑S�Ăɉ񓚂�������
                        If answer_num = target_range Then
                            
                            ' �t���O��L���ɂ���
                            ws_calculation.Cells(indata_count, 1) = 1
                            
                        Else
                            
                            ' �L���o�Ȃ��ꍇ�̓N���A���s��
                            ws_calculation.Cells(indata_count, 1) = ""
                            
                        End If
            
                    ' �󗓂�������
                    Case Else
            
                End Select
            
            Next indata_count
            
            ' �v�Z�V�[�g���A�N�e�B�u�ɕύX
            ws_calculation.Activate
        
            ' �ʏ폈��
            target_address = ws_calculation.Cells(START_ROW_INDATA, 1).Address
            
            ' �Ώۂ͈̔͂��R�s�[���ē��̓f�[�^�ɓ\��t����
            ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
            Selection.Copy
            
            ' ���̓f�[�^���A�N�e�B�u�ɕύX
            wsp_indata.Activate
            
            ' �������݈ʒu�̃A�h���X���擾
            target_address = wsp_indata.Cells(START_ROW_INDATA, q_data(ws_calculation.Cells(QCODE_P_ROW, 1).Value).data_column).Address
            
            ' �R�s�[�����f�[�^��\��t��
            wsp_indata.Range(target_address).PasteSpecial (xlPasteValues)
            
            ' �쐬�����L�[�����̏����Ɋ܂܂�Ă���ꍇ
            On Error Resume Next
            match_column = WorksheetFunction.Match(ws_calculation.Range("A4"), ws_calculation.Rows(3), 0)
            
            ' �L�[����v�����ꍇ
            If match_column Then
                
                ' �擪�A�h���X���擾
                target_address = ws_calculation.Cells(START_ROW_INDATA, match_column).Address
                
                ' �R�s�[�����f�[�^��\��t��
                ws_calculation.Range(target_address).PasteSpecial (xlPasteValues)
            
            End If
            
            On Error GoTo 0
            
        ' �n�_�A�I�_�������͈͂̌v�Z
        Else
        
            ' �J�E���^�[�𑝉�����
            target_range = target_range + 1
        
        End If
    
    Next process_count
    
    ' �I�u�W�F�N�g�����
    wb_calculation.Close SaveChanges:=False
    
End Sub



'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2017.06.26  '
' �f�[�^�N���A�p�v���V�[�W��                                                                 '
' �����P Long�^      Select�pQCODE��ԍ�                                                           '
' �����Q Long�^      Select�pValue���e                                                             '
' �����R Long�^      QCODE�����ԍ��i�[�p                                                           '
' �����S WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����T Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub select_clear(ByVal select_row As Long, ByVal select_value As Long, ByVal qcode_count As Long, _
ByVal wsp_indata As Worksheet, ByVal indata_maxrow As Long)

    Dim indata_count As Long        ' ���̓��R�[�h�J�E���g�p�ϐ�
    Dim select_address As String    ' �Z���N�g�A�h���X�i�[�p�ϐ��i�����񓚂̎��̂ݎg�p�j
    
    Dim clear_address As String     ' �N���A�͈͊i�[�p�ϐ�

    ' �N���A�����擾
    Select Case q_data(qcode_count).q_format
    
        ' �P���
        Case "S", "F", "O"
        
            ' �N���A�͈͂��擾
            clear_address = wsp_indata.Cells(6, q_data(qcode_count).data_column).Address
    
        ' ������
        Case "M", "L", "LM", "LA", "LC"
            
            ' �N���A�͈͂��擾
            clear_address = wsp_indata.Cells(6, q_data(qcode_count).data_column).Resize(, q_data(qcode_count).ct_count).Address
                
        ' �C���M�����[��Format
        Case Else

            clear_address = ""

    End Select

    ' �N���A�͈͂��擾���Ă���ꍇ
    If clear_address <> "" Then
        ' �Z���N�g�����̃t�H�[�}�b�g�𔻒�
        Select Case q_data(select_row).q_format

            ' �P���W�v
            Case "S", "F", "O"
            
                ' 0�J�e�S���[�����̂܂܏���
                For indata_count = START_ROW_INDATA To indata_maxrow
            
                    ' �񓚓��e����v������
                    If wsp_indata.Cells(indata_count, q_data(select_row).data_column).Value = select_value Then
                
                        ' �������s��Ȃ�
                
                    ' �񓚓��e����v���Ȃ�������
                    Else
                    
                        ' �͈͂��N���A
                        wsp_indata.Range(clear_address).Offset(indata_count - 6).ClearContents
                
                    End If
            
                Next indata_count
            
            ' �����񓚊֘A
            Case "M", "L", "LM", "LA", "LC"
        
                ' ���̓f�[�^�𔻒�
                For indata_count = START_ROW_INDATA To indata_maxrow
                    
                    ' �w��̋L���l�����鎞
                    If Val(wsp_indata.Cells(indata_count, q_data(select_row).data_column).Offset(, select_value).Value) > 0 Then
                    ' �����s��Ȃ�
                    ' �w��̋L���l���Ȃ���
                    Else
                            ' ���͒l���N���A
                        wsp_indata.Range(clear_address).Offset(indata_count - 6, 0).ClearContents
                    End If
            
                Next indata_count

            ' �C���M�����[��Format
            Case Else

        End Select

    End If

End Sub




'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2020.03.30  '
' �������H�p�v���V�[�W��                                                                           '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub amplification_data(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim str_address() As String     ' �����A�h���X�i�[�p�z��
    Dim str_qcode() As Long         ' ����QCODE�i�[�p�z��
    Dim str_outqcode() As String    ' �o��QCODE�i�[�p�z��
    Dim target_coderow As Long      ' �Ώ�QCODE��ԍ��ꎞ�i�[�p�ϐ�
    
    Dim column_count As Long        ' ������ԍ��i�[�p�ϐ�
    Dim column_end As Long          ' ��I�[�ԍ��i�[�p�ϐ�
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim rows_count As Long          ' ���͈ʒu�i�[�p�ϐ�

    Dim str_ctdata As Variant       ' �J�e�S���C�Y�e�[�u�����i�[�p�z��ϐ�
    
    'Dim writing_column As Long      ' �������݈ʒu�i�[�p�ϐ�
    'Dim target_data As Double       ' ���̓f�[�^����p�ϐ�
    'Dim categorize_count As Long    ' ����J�E���g�p�ϐ�
    'Dim table_max As Long           ' �I�[�e�[�u���ԍ��i�[�p�ϐ�
    'Dim ma_count As Long            ' MA�J�e�S���[�����擾
    
    ' ���̓f�[�^MA����p�ϐ��Q
    Dim str_madata As Variant       ' MA�񓚃f�[�^�i�[�p�ϐ�
    Dim ma_address As String        ' MA�񓚃f�[�^�A�h���X�i�[�p�ϐ�
    Dim maindata_count As Long      ' MA�f�[�^�J�E���g�ϐ�
    
    Dim prosessing_flg As Boolean   ' �w�b�_���H����p�t���O

    Dim pmax_number As Double       ' �e�[�u���ő�l�i�[�p�ϐ�
    Dim pmin_number As Double       ' �e�[�u���ŏ��l�i�[�p�ϐ�

    Dim str_count As Long           ' str_ctdata�J�E���g�p�ϐ�
    
    Dim identity_count As Long      ' �����NO�J�E���g�p�ϐ�
    Dim identity_max As Long        ' �����NO�ő吔�i�[�p�ϐ�
    
    Dim stray_ws As Worksheet       ' ���J�e�S���C�Y�f�[�^���O�o�͗p�I�u�W�F�N�g�ϐ�
    Dim ct_flg As Boolean           ' ���J�e�S���C�Y�f�[�^����p�t���O
    
    Dim amp_qcode As String         ' �����ݖ▼�i�[�p�ϐ�
    
    Dim table_count As Long         ' �e�[�u���A�h���X���J�E���g�p�ϐ�
    Dim table_max As Long           ' �ő�e�[�u�����i�[�p�ϐ�
    
    Dim target_table As Variant     ' �����e�[�u�����i�[�p�ϐ�
    Dim table_count_y As Long       ' �i�[�e�[�u�����Q�Ɨp�ϐ��i�c���j
    Dim table_count_x As Long       ' �i�[�e�[�u�����Q�Ɨp�ϐ��i�����j
    
    
    'Dim log_rows As Long            ' ���O�o�͈ʒu�i�[�p�ϐ�
    
    ' �����ݒ�
    ReDim str_address(300)
    ReDim str_qcode(300)
    ReDim str_outqcode(300)
    
    ' ��ʂւ̕\�����I���ɂ���
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "�@�J�e�S���C�Y���H�������v�Z��..."
    
    ' ��ʂւ̕\�����I�t�ɂ���
    'Application.ScreenUpdating = False
    
    
    ' �I�[��ԍ��̎擾 20170502 START_ROW ��6�ɕύX
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' �J�E���g����������
    table_count = 1
    
    ' ���J�e�S���C�Y���i�[�p�V�[�g�ǉ�
    'Set stray_ws = error_tb.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    
    'stray_ws.Name = "���J�e�S���C�Y���X�g"
    'stray_ws.Range("A1").Value = "SampleNo"
    'stray_ws.Range("B1").Value = "QCODE"
    'stray_ws.Range("C1").Value = "MA_CT"
    'stray_ws.Range("D1").Value = "�G���[���e"
    'stray_ws.Range("E1").Value = "�񓚓��e"
    'stray_ws.Range("F1").Value = "�C�����e"
    'stray_ws.Range("G1").Value = "����"

    ' �������R�[�h�����擾
    For column_count = 12 To column_end
    
        ' �A�X�^���X�N���������Ƃ����𐔂���
        ' ���A�X�^���X�N�Ɗ����b�s���Œ�ʒu���Ώېݖ�ɋL�����聕Skip�t���O����
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 2, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value = "" Then
            
            ' �A�h���X���擾
            str_address(table_count) = ws_process.Cells(START_TABLE_DATA, column_count).Address
            table_count = table_count + 1
            
            ' ���ڐ����J�E���g
            identity_count = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row - 5
            
            ' �J�E���g�������ڐ��������Ƃ������ꍇ�A���ڐ����擾
            If identity_max < identity_count Then
                identity_max = identity_count
            End If
            
            'str_address = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row
            
        End If
    
    Next column_count
    
    ' �ő�e�[�u�������擾
    table_max = table_count
    
    ' ���̓f�[�^���ׂĂɁA�S�e�[�u�����̏������s��
    ' �@�ˁ@�񓚂𑝕��A�e�[�u�������̏������s��
    For indata_count = START_ROW_INDATA To indata_maxrow
    
        ' 20200402 �ǋL
        'wsp_indata.Rows (indata_count)
    
        ' �P���R�[�h���R�s�[
        wsp_indata.Rows(indata_count).Copy
        ' �\��t����̍��W���擾
        rows_count = wsp_indata.Cells(Rows.Count, 1).End(xlUp).Row + 1
        ' �f�[�^�̑���
        wsp_indata.Rows(rows_count & ":" & (rows_count + identity_count - 1)).PasteSpecial
        
    Next indata_count
    
    ' �e�[�u�������������s��
    For table_count = 1 To table_max
        
        ' �e�[�u���������ׂĊi�[
        target_table = ws_process.Range(str_address(table_count)).Resize(299, 12).Value
        
        ' �c�����[�v
        For table_count_y = 1 To identity_max
        
            ' Skip�t���O�������Ă��Ȃ��ꍇ
            If target_table(table_count_y, 1) = "" Then
        
                ' �������[�v
                For table_count_x = 3 To identity_max
            
                    ' �Ώۖ�m�����N���Aor�������
                    Qcode_Match (ws_process.Cells(table_count_y, table_count_x).Value)
                    
            
                Next table_count_x
        
            End If
        
        Next table_count_y
        
        
        
    Next table_count
    
End Sub



'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2018.06.14  '
' ���O�i�[�p�v���V�[�W��                                                                           '
' �����P String�^    ���H���e�f�[�^                                                                '
' �����Q String�^    QCODE1�f�[�^                                                                  '
' �����R String�^    QCODE2�f�[�^                                                                  '
' �����S String�^    �������e�f�[�^                                                                '
' �����T WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub print_log(ByVal log_data1 As String, ByVal log_data2 As String, ByVal log_data3 As String, _
ByVal log_data4 As String, ByVal ws_logs As Worksheet)

    Dim row_data As Long    ' �ŏI�s�i�[�p�ϐ�

    ' �ŏI�s�擾
    row_data = ws_logs.Cells(Rows.Count, 1).End(xlUp).Row

    ' SEQ
    ws_logs.Cells(row_data + 1, 1) = Format(row_data)
    
    ' ���̑��f�[�^�̏o��
    ws_logs.Cells(row_data + 1, 2) = log_data1
    ws_logs.Cells(row_data + 1, 3) = log_data2
    ws_logs.Cells(row_data + 1, 4) = log_data3
    ws_logs.Cells(row_data + 1, 5) = log_data4
    
End Sub

'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2018.06.29  '
' ���O�i�[�p�v���V�[�W���i�J�e�S���C�Y�p�j                                                         '
' �����P Long�^      ���R�[�h���                                                                  '
' �����Q String�^    QCODE1�f�[�^                                                                  '
' �����R Long�^      MA_CT                                                                         '
' �����S String�^    �񓚓��e�f�[�^                                                                '
' �����T WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub print_log2(ByVal sample_row As Long, ByVal qcode_data As String, ByVal ma_ct As Long, _
ByVal input_data As String, ByVal ws_logs As Worksheet)

    Dim row_data As Long    ' �ŏI�s�i�[�p�ϐ�

    ' �ŏI�s�擾
    row_data = ws_logs.Cells(Rows.Count, 1).End(xlUp).Row

    ' SEQ
    'ws_logs.Cells(row_data + 1, 1) = Format(row_data)
    
    ' ���̑��f�[�^�̏o��
    'ws_logs.Cells(row_data + 1, 2) = log_data1
    'ws_logs.Cells(row_data + 1, 3) = log_data2
    'ws_logs.Cells(row_data + 1, 4) = log_data3
    'ws_logs.Cells(row_data + 1, 5) = log_data4
    
End Sub


'--------------------------------------------------------------------------------------------------'
' �쐬��  ���R��                                                               �쐬��  2017.05.08  '
' �Z���N�g�t���O���H�p�v���V�[�W��                                                                 '
' �����P WorkSheet�^ �J�e�S���C�Y�����w���V�[�g                                                    '
' �����Q WorkSheet�^ ���̓f�[�^�V�[�g                                                              '
' �����R Long�^      ���̓f�[�^�I�[�ԍ�                                                            '
' �����S WorkSheet�^ ���H���O�o�̓V�[�g                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg3(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' ���H�񐔃J�E���g�p�ϐ�
    Dim process_max As Long         ' �ő���H�񐔊i�[�p�ϐ�
    
    Dim qcode1_row As Long          ' ��r�ݖ�(�q) ROW�i�[�p�ϐ�
    Dim qcode2_row As Long          ' �t�Z�b�g�Ώېݖ�(�e) ROW�i�[�p�ϐ�
    Dim qcode3_row As Long          ' �G���g���[�G���A�I�[�i�[�p�ϐ�
    
    Dim input_word As String        ' �t�Z�b�g�pInputWord�i�[�p������ϐ�
    Dim process_flg As Boolean      ' ��������p�t���O
    
    Dim indata_count As Long        ' ���̓f�[�^�����ʒu�i�[�p�ϐ�
    Dim work_maxcol As Long         ' �w���񓚏I�[�ʒu
    
    Dim str_ct1 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct2 As Variant          ' �J�e�S���[��r�p�z��i���j
    Dim str_ct3 As Variant          ' �J�e�S���[��r�p�z��i��+���j
    
    Dim ct_count As Long            ' �z����e�J�E���g�p�ϐ�
    
'    Dim ct1_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct2_count As Long           ' ���񓚐��i�[�p�ϐ�
'    Dim ct3_count As Long           ' ��+���񓚐��i�[�p�ϐ�
    
    Dim work1_flg As Boolean        ' �������i�[�p�t���O
    Dim work2_flg As Boolean        ' �������i�[�p�t���O
    
    Dim target_address As String    ' QCODE1�A�h���X�i�[�p�ϐ�
    
    Dim processing_flg As Boolean   ' Function�߂�l�i�[�p�ϐ�
    
    Dim wb_calculation As Workbook  ' ���Ԍv�Z�p�u�b�N���i�[�p�ϐ�
    Dim ws_calculation As Worksheet ' ���Ԍv�Z�p�V�[�g���i�[�p�ϐ�
    
    Dim start_num As Double         ' �J�n�ԍ��i�[�p�ϐ�
    Dim end_num As Double           ' �I�[�ԍ��i�[�p�ϐ�
    Dim connect_data As String      ' �ڑ����i�[�p�ϐ�
    Dim target_range As Long        ' �Z���N�g�͈͎擾�p�ϐ�
    Dim answer_num As Long          ' �񓚐��i�[�p�ϐ�
    Dim match_data As Long          ' �쐬QCODE�s�ԍ��i�[�p�ϐ�
    
    Dim work_flg As Long            ' �������������t���O
    Dim work_address As String      ' ���������͈͊i�[�p�ϐ�
    Dim match_column As Long
    
' �e�w���P�ꖈ�ɏ����𖞂����Ă��邩�𔻒肵�ĕʃu�b�N�ɓf���o��----------------------------------------------
    
    ' ���H���������擾
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' �v�Z�p�u�b�N���쐬
    Set wb_calculation = Workbooks.Add
    Set ws_calculation = wb_calculation.Worksheets(1)
    
    ' �ő���H�񐔕��������s��
    For process_count = START_ROW To process_max
    
        ' �X�L�b�v�t���O�����͂���Ă��Ȃ���
        If Len(ws_process.Cells(process_count, SKIP_FLG).Value) = 0 Then
        
            ' �Q�Ɛݖ��QCODE�̗�ԍ����擾
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            
            ' �����f�[�^�ł͂Ȃ��EQCODE2�����񓚂ł͂Ȃ�
            If ws_process.Cells(process_count, QCODE1_DATA4).Value <> "" Then
                qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4).Value)
            ' QCODE2�����񓚂̎��A���������t���O���L���̏ꍇ�͎�O�̏o�͐ݒ���R�s�[
            ElseIf process_count <> START_ROW And ws_process.Cells(process_count, QCODE1_DATA4).Value = "" Then
                qcode2_row = Qcode_Match(ws_process.Cells(process_count - 1, QCODE1_DATA4).Value)
                ws_process.Cells(process_count, QCODE1_DATA4).Value = ws_process.Cells(process_count - 1, QCODE1_DATA4).Value
            
            Else
            ' �����f�[�^�������͏o�͐ݖ�ԍ������񓚂̎�
                
            End If
            
            qcode3_row = Qcode_Match("*���H��")
            
            ' �o�͐ݖ�ԍ��𔻒肵�A�w�b�_�[���쐬����
            If q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' �ʏ폈��
            
            Else
            
                ' �w�b�_���쐬����
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' �V�����ݒ肵���G���A��q_data�ɃJ�����Ƃ��Đݒ肷��
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' �ŏ��l�A�ő�l����̓f�[�^�̃w�b�_�Ɋi�[
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' �v�Z�V�[�g�ɓ���QCODE��������
            ws_calculation.Cells(START_ROW - 3, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1)
            
            ' �v�Z�V�[�g�ɏo��QCODE��������
            ws_calculation.Cells(START_ROW - 2, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA4)
            
            ' �v�Z�V�[�g�ɏo�͐�QCODE��ԍ���������
            ws_calculation.Cells(START_ROW - 1, process_count - QCODE_P_COLUMN).Value = qcode2_row
            
            ' �v�Z�V�[�g�ɏW�v������������
            ws_calculation.Cells(START_ROW, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA5)

            ' �J�n�ԍ��ƏI�[�ԍ���ϐ��ɕێ�������i�V���[�g�J�b�g�ȊO�j
            If ws_process.Cells(process_count, QCODE1_DATA1).Value <> "*" And _
            ws_process.Cells(process_count, QCODE1_DATA1).Value <> "_" Then
                start_num = ws_process.Cells(process_count, QCODE1_DATA1).Value
                end_num = ws_process.Cells(process_count, QCODE1_DATA3).Value
                
                ' �ŏ��l�E�ő�l��ݒ�
                ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = start_num
                ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = end_num
            
            ' �V���[�g�J�b�g�g�p���͏�����
            Else
                start_num = 0
                end_num = 0
                
                If ws_process.Cells(process_count, QCODE1_DATA1).Value = "*" Then
                
                    ' �ŏ��l�E�ő�l��ݒ�
                    ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = "*"
                    ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = "*"
            
                ElseIf ws_process.Cells(process_count, QCODE1_DATA1).Value = "_" Then
                
                    ' �ŏ��l�E�ő�l��ݒ�
                    ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = "_"
                    ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = "_"
                
                End If
            
            End If

            ' ���̓f�[�^�S�Ăɏ������s��
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' �Ώېݖ�̃t�H�[�}�b�g�ɂ�蔻�������ύX
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' �P��񓚂̏ꍇ
                    Case "S", "R", "H"
                
                        ' ��������񓚂����鎞
                        If ws_process.Cells(process_count, QCODE1_DATA1).Value = "*" Then
                            
                            If Len(Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value)) <> 0 Then
                            
                                ' �t���O�𗧂Ă�
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            End If
                        
                        ' �����񓚂��Ȃ���
                        ElseIf ws_process.Cells(process_count, QCODE1_DATA1).Value = "_" Then
                            
                            If Len(Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value)) = 0 Then
                            
                                ' �t���O�𗧂Ă�
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            End If
                            
                        ' �w��͈͂ɉ񓚂����鎞
                        ElseIf wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value >= start_num And _
                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <= end_num Then
                            
                            ' �t���O�𗧂Ă�
                            ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            
                        End If
                    
                    ' �����񓚂̏ꍇ
                    Case "L", "M"
                    
                        ' �擪�A�h���X���擾
                        target_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            
                            ' �Ώۂ͈̔͂̉񓚂��m�F����
                            If WorksheetFunction.Sum(wsp_indata.Range(target_address) _
                            .Offset(0, start_num - 1).Resize(, end_num + 1 - start_num)) <> 0 Then
                        
                                ' �t���O�𗧂Ă�
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                            End If

                    Case Else
                    
                    End Select
                    
            Next indata_count
        
        End If
    
    Next process_count
    
' �ʃu�b�N�ɓf���o����������𐮌`���Z���N�g�t���O���쐬����---------------------------------------------
    
    ' �v�Z�V�[�g���A�N�e�B�u�ɕύX
    ws_calculation.Activate
    
    ' �ő���H�񐔕��������s��
    For process_count = (START_ROW - QCODE_P_COLUMN) To (process_max - QCODE_P_COLUMN)
        
        ' ���������t���O�𖢎g�p
        If ws_calculation.Cells(6, process_count).Value = "" Then
        
            ' �J�n�ԍ��E�I�[�ԍ���������Ă��Ȃ���
            If ws_calculation.Cells(1, process_count).Value <> "" Then
        
                ' �v�Z�V�[�g���A�N�e�B�u�ɕύX
                ws_calculation.Activate
        
                ' �ʏ폈��
                target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            
                ' �Ώۂ͈̔͂��R�s�[���ē��̓f�[�^�ɓ\��t����
                ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
                Selection.Copy
            
                ' ���̓f�[�^���A�N�e�B�u�ɕύX
                wsp_indata.Activate
            
                ' �������݈ʒu�̃A�h���X���擾
                target_address = wsp_indata.Cells(START_ROW_INDATA, _
                q_data(ws_calculation.Cells(QCODE_P_ROW, process_count).Value).data_column).Address
            
                ' �R�s�[�����f�[�^��\��t��(�󔒂𖳎����Ē���t����)
                wsp_indata.Range(target_address).PasteSpecial xlPasteValues, SkipBlanks:=True
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then
                
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "�`" & _
                    ws_calculation.Cells(2, process_count).Value & "�̉񓚓��e���o�́B", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�L���񓚂ŏo�́B", ws_logs)
                ElseIf ws_calculation.Cells(1, process_count).Value = "_" Then
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�����񓚂ŏo�́B", ws_logs)
                End If
                
            
            ' �G���[�t���O�������Ă��鎞
            Else
            
                ' �G���[�o��
                Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                ws_calculation.Cells(3, process_count - 1).Value & "�Ƃ́u" & _
                ws_calculation.Cells(START_ROW, process_count - 1).Value & _
                "�v�����̎w���ɑ΂��A�قȂ�o�͐ݖ�ԍ����ݒ肳��Ă��܂��B", ws_logs)
            
            End If
            
        ' �������������鎞�A�o�͐ݒ�ԍ�������̎��i�����i�s�j
        ElseIf ws_calculation.Cells(4, process_count).Value = ws_calculation.Cells(4, process_count + 1).Value Then
        
            
            ' ���������̓��e�����i�`�m�c�����j
            If Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "an" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "An" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "AN" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "����" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "�`�m" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "�`��" Then
            
                work_flg = 1
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then
                
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "�`" & _
                    ws_calculation.Cells(2, process_count).Value & "�̉񓚓��e��" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃ`�m�c�����ŏo�͐ݒ�B", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�L���񓚂�" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃ`�m�c�����ŏo�͐ݒ�B", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "_" Then
                    
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�����񓚂�" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃ`�m�c�����ŏo�͐ݒ�B", ws_logs)
                
                Else
                
                End If
                
            
            ' ���������̓��e�����i�n�q�����j
            ElseIf Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "or" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "Or" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "OR" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "����" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "�n�q" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "�n��" Then
            
                work_flg = 2
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then

                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "�`" & _
                    ws_calculation.Cells(2, process_count).Value & "�̉񓚓��e��" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃn�q�����ŏo�͐ݒ�B", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                    
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�L���񓚂�" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃn�q�����ŏo�͐ݒ�B", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                
                    ' �������e���o��
                    Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "�����񓚂�" & _
                    ws_calculation.Cells(3, process_count + 1) & "�Ƃn�q�����ŏo�͐ݒ�B", ws_logs)
                
                Else
                
                End If
                
            ' �w���ȊO�̏�񂪓����Ă���ꍇ
            Else
            
                work_flg = 0
            
            End If
            
            ' ���̓f�[�^�S�Ăɏ������s��
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' �A�h���X�ňʒu�����擾�i���[�N�V�[�g�t�@���N�V�������g�p���邽�߁j
                work_address = ws_calculation.Cells(indata_count, process_count).Address
                
                ' �`�m�c�����̎�
                If work_flg = 1 Then
                    
                    ' �t���O�̍��v��2�ȏ�̎�
                    If WorksheetFunction.Sum(ws_calculation.Range(work_address).Resize(, 2).Value) = 2 Then
                        ws_calculation.Range(work_address).Offset(, 1).Value = 1
                    Else
                        ws_calculation.Range(work_address).Offset(, 1).Value = ""
                    End If
                
                ' �n�q�����̎�
                ElseIf work_flg = 2 Then
                    
                    ' �t���O�̍��v��2�ȏ�̎�
                    If WorksheetFunction.Sum(ws_calculation.Range(work_address).Resize(, 2).Value) > 0 Then
                        ws_calculation.Range(work_address).Offset(, 1).Value = 1
                    Else
                        ws_calculation.Range(work_address).Offset(, 1).Value = ""
                    End If
                
                ' ��L�ȊO�̎�
                Else
                
                    ' ��O�Ŕ�������Ă���̂ł����ɓ��邱�Ƃ͂Ȃ��\��
                    ' �\���Ƃ��Ă̓G���g���[�G���A�ւ̎w���̎��ɒʂ�\��������
                
                End If

            Next indata_count

            
        ' ���������̐ݒ肪����ɂ��ւ�炸�o�͐ݖ�ԍ����قȂ鎞
        Else
        
            ' �G���[�o��
            Call print_log("�Z���N�g�t���O����", ws_calculation.Cells(3, process_count), _
            ws_calculation.Cells(START_ROW - 2, process_count).Value, _
            ws_calculation.Cells(3, process_count + 1).Value & "�Ƃ́u" & _
            ws_calculation.Cells(START_ROW, process_count).Value & _
            "�v�����̎w���ɑ΂��A�قȂ�o�͐ݖ�ԍ����ݒ肳��Ă��܂��B", ws_logs)
            
            ' ��������p�̊J�n�ԍ��E�I�[�ԍ�������
            ws_calculation.Cells(START_ROW - 4, process_count).Value = ""
            ws_calculation.Cells(START_ROW - 5, process_count).Value = ""
            
            ws_calculation.Cells(START_ROW - 4, process_count + 1).Value = ""
            ws_calculation.Cells(START_ROW - 5, process_count + 1).Value = ""
            
        End If
    
    Next process_count
    
    ' �I�u�W�F�N�g�����
    wb_calculation.Close SaveChanges:=False
    
End Sub



