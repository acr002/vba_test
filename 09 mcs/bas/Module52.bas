Attribute VB_Name = "Module52"
Option Explicit
    Public wb_tabinst As Workbook
    Public ws_tabinst As Worksheet
    
    Public wb_summary As Workbook
    Public ws_summary0 As Worksheet
    Public ws_summary1 As Worksheet
    Public ws_summary2 As Worksheet
    Public ws_summary3 As Worksheet
    
    Public tabinst_fn As String
    Public summary_fn As String
    Public summary_fd As String
    
    Dim c_cnt As Long
    Dim outdata_row As Long             ' �W�v�f�[�^�̍ŏI�s
    Dim outdata_col As Long             ' �W�v�f�[�^�̍ŏI��
    
    Dim a_index, f_index As Long        ' �\���E�\���̐ݒ���̃C���f�b�N�X
    Dim ra_index As Long                ' �����ݖ�̃C���f�b�N�X
    
    ' �\���E�\���E�W�v�ݒ�ɁA�Z���N�g�̐ݒ肪���邩�Ȃ����̔���p�ł��B
    ' �f�[�^���Z���N�g�����𖞂����Ă��邩�́A16384��ڂŔ��肵�܂��B
    Dim select_flg As Integer           ' �Z���N�g�ݒ�̗L���t���O
    
    ' 2019.10.10 - �E�G�C�g�W�v�֘A�ǉ�
    Dim weight_flg As String            ' �E�G�C�g�W�v�L���̃t���O�i�Ȃ��A����j
    Dim weight_col As Long              ' �␳�l�̗�
    Dim w_index As Long                 ' �␳�l�̃C���f�b�N�X
    
    Dim se_index As Long                ' �W�v�ݒ�Z���N�g�̃C���f�b�N�X
    Dim as1_index As Long               ' �\���Z���N�g�̃C���f�b�N�X
    Dim as2_index As Long
    Dim as3_index As Long
    Dim fs1_index As Long               ' �\���Z���N�g�̃C���f�b�N�X
    Dim fs2_index As Long
    Dim fs3_index As Long
    
    Dim ama_cnt As Long                 ' �\���̂l�`�J�e�S���[���i0:SA�A0�ȊO:MA�j
    Dim fma_cnt As Long                 ' �\���̂l�`�J�e�S���[���i0:SA�A0�ȊO:MA�j
    
    Dim se_cnt As Long                  ' �W�v�ݒ�Z���N�g�̂l�`�J�e�S���[���i0:SA�A0�ȊO:MA�j
    Dim as1_cnt As Long                 ' �\���Z���N�g�̂l�`�J�e�S���[���i0:SA�A0�ȊO:MA�j
    Dim as2_cnt As Long
    Dim as3_cnt As Long
    Dim fs1_cnt As Long                 ' �\���Z���N�g�̂l�`�J�e�S���[���i0:SA�A0�ȊO:MA�j
    Dim fs2_cnt As Long
    Dim fs3_cnt As Long
    
    Dim se_msg As String                ' �W�v�ݒ�Z���N�g�̃��b�Z�[�W
    Dim as1_msg As String               ' �\���Z���N�g�̃��b�Z�[�W
    Dim as2_msg As String
    Dim as3_msg As String
    Dim fs1_msg As String               ' �\���Z���N�g�̃��b�Z�[�W
    Dim fs2_msg As String
    Dim fs3_msg As String
    
    Dim sum_row As Long                 ' �W�v�\�i�m���\�j�̍s��
    Dim sum_col As Long                 ' �W�v�\�i�m���\�j�̍s��
    
    Dim div_row As Long                 ' �W�v�\�i�m�\�A���\�j�̍s��
    Dim div_col As Long                 ' �W�v�\�i�m�\�A���\�j�̍s��
    
    Public Type cross_data
        hyo_num As String               ' �\���i�[�p������ϐ�
        f_code As String                ' �\��QCODE�i�[�p������ϐ�
        a_code As String                ' �\��QCODE�i�[�p������ϐ�
        r_code As String                ' �����ݖ�QCODE�i�[�p������ϐ�
        fna_flg As String               ' �\���\���t���O�i�[�p������ϐ�
        ana_flg As String               ' �\��NA�t���O�i�[�p������ϐ�
        bosu_flg As String              ' �W�v�ꐔ�t���O�i�[�p�ϐ�
    
        sum_flg As String               ' �����ݖ�o�͎w���E���v�t���O�i�[�p������ϐ�
        ave_flg As String               ' �����ݖ�o�͎w���E���σt���O�i�[�p������ϐ�
        sd_flg As String                ' �����ݖ�o�͎w���E�W���΍��t���O�i�[�p������ϐ�
        min_flg As String               ' �����ݖ�o�͎w���E�ŏ��l�t���O�i�[�p������ϐ�
        q1_flg As String                ' �����ݖ�o�͎w���E��P�l���ʃt���O�i�[�p������ϐ�
        med_flg As String               ' �����ݖ�o�͎w���E�����l�t���O�i�[�p������ϐ�
        q3_flg As String                ' �����ݖ�o�͎w���E��R�l���ʃt���O�i�[�p������ϐ�
        max_flg As String               ' �����ݖ�o�͎w���E�ő�l�t���O�i�[�p������ϐ�
        mod_flg As String               ' �����ݖ�o�͎w���E�ŕp�l�t���O�i�[�p������ϐ�
    
        sel_code As String              ' �Z���N�g�����EQCODE
        sel_value As Integer            ' �Z���N�g�����E�l
        
        ken_flg As String               ' �\���I�v�V�����E�������t���O
        yuko_flg As String              ' �\���I�v�V�����E�L���񓚃t���O
        nobe_flg As String              ' �\���I�v�V�����E���׉�
    
        top1_flg As String              ' TOP1�E�}�[�L���O�t���O
        sort_flg As String              ' CT�\�[�g�E�~���t���O
        exct_flg As String              ' CT�\�[�g�E���OCT�t���O
        graph_flg As String             ' �O���t�E�쐬�t���O
    End Type
    
    Dim c_data() As cross_data          ' �\�����̏W�v�w����S�Ď擾

Sub Summarydata_Creation()
    Dim yensign_pos As Long
    Dim r_code As Integer
'2018/05/23 - �ǋL ==========================
    Dim crs_tab() As String
    Dim crs_file As String
    Dim crs_cnt As Long
    Dim i_cnt As Long
    Dim fn_cnt As Long
'--------------------------------------------------------------------------------------------------'
'�@�W�v�T�}���[�f�[�^�̍쐬�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.05.22�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
'�y�Y�ꌾ�z2017.05.10�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'�@�T�}���[�t�@�C���쐬�O�Ƀ��W�b�N�`�F�b�N��g�ݍ��ނ��������܂������A�W�v�ݒ�t�@�C������������@'
'�@�ꍇ�Ȃǂ̃P�[�X�ŁA�`�F�b�N�P��A�W�v��������̂Ƃ��ɖ��񓯂��`�F�b�N������Ǝ��Ԃ�������̂ŁA'
'�@���W���[���͓Ɨ��������̂Ƃ��܂��B�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    Application.StatusBar = "�W�v�T�}���[�f�[�^�̍쐬 ��������ƒ�..."
    
    Open file_path & "\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���NG.xlsx" For Append As #1
    Close #1
    If Err.Number > 0 Then
        MsgBox "�ݒ��ʃG���[�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���NG.xlsx�n���J����Ă��܂��B" _
        & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "�ݒ��ʂ̓��͏��ɃG���[������\��������܂��B" _
        & vbCrLf & "�G���[�̓��e���m�F���ām" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���NG.xlsx�n����Ă���A" _
        & vbCrLf & "�Ď��s���Ă��������B" _
        , vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    Open file_path & "\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx" For Append As #1
    Close #1
    If Err.Number > 0 Then
        MsgBox "���W�b�N�`�F�b�N�̃G���[���O�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx�n���J����Ă��܂��B" _
        & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "���W�b�N�`�F�b�N���s�����f�[�^�ɃG���[������\��������܂��B" _
        & vbCrLf & "�G���[�̓��e���m�F���ām" & ws_mainmenu.Cells(gcode_row, gcode_col) & "err.xlsx�n����Ă���A" _
        & vbCrLf & "�Ď��s���Ă��������B" _
        , vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ' ���O�`�F�b�N�̃��b�Z�[�W
    wb.Activate
    ws_mainmenu.Select
    r_code = MsgBox("���O�ɏW�v����t�@�C���̃��W�b�N�`�F�b�N���s���܂������B" _
    & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "���O�Ƀ��W�b�N�`�F�b�N���s���Ώۃt�@�C���́A" _
    & vbCrLf & "�W�v�ݒ�t�@�C���́m�W�v����t�@�C�����n�̍��ڂ�" & vbCrLf & "�w�肵���t�@�C���ɂȂ�܂��B", _
    vbYesNo + vbQuestion, "MCS 2020 - Summarydata_Creation")
    If r_code = vbNo Then
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ChDrive file_path
    ChDir file_path & "\3_FD"
    
    ' CRS�t�H���_����xlsx�`���̃t�@�C�������J�E���g
    crs_cnt = 0
    crs_file = Dir(file_path & "\3_FD\CRS\*.xlsx")
    Do Until crs_file = ""
        DoEvents
        crs_cnt = crs_cnt + 1
        crs_file = Dir()
    Loop
    
    ' CRS�t�H���_����xlsx�`���̃t�@�C������z��ɃZ�b�g
    ReDim crs_tab(crs_cnt)
    crs_file = Dir(file_path & "\3_FD\CRS\*.xlsx")
    For fn_cnt = 1 To crs_cnt
        DoEvents
        crs_tab(fn_cnt) = crs_file
        crs_file = Dir()
    Next fn_cnt
    fn_cnt = crs_cnt
    
' �W�v�T�}���[�f�[�^������쐬����
    If crs_cnt > 0 Then
        r_code = MsgBox("FD�t�H���_���ɁmCRS�n�t�H���_������܂��B" _
         & vbCrLf & "CRS�t�H���_���ɂ���" & fn_cnt & "�̏W�v�ݒ�t�@�C�����g�p���āA" _
         & vbCrLf & "�W�v�T�}���[�f�[�^���ꊇ�쐬���܂����B" _
         & vbCrLf & vbCrLf & "�yTIPS�z" & vbCrLf & "CRS�t�H���_���́mxlsx�`���n�̃t�@�C������\�����Ă��܂��B" _
         & vbCrLf & "�u�͂��v�@�� �W�v�T�}���[�f�[�^���ꊇ�쐬" & vbCrLf & "�u�������v�� �W�v�ݒ�t�@�C����I�����Ă���쐬", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Summarydata_Creation")
        If r_code = vbYes Then
            crs_cnt = 1
            For i_cnt = 1 To fn_cnt
                DoEvents
                tabinst_fn = crs_tab(i_cnt)
                wb.Activate
                ws_mainmenu.Select
                If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\CRS\" & tabinst_fn) <> "" Then
                    Workbooks.Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & _
                     ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\CRS\" & tabinst_fn
                    Set wb_tabinst = Workbooks(tabinst_fn)
                    Set ws_tabinst = wb_tabinst.ActiveSheet
                    If ws_tabinst.Cells(2, 4) = "" Then
                        MsgBox "�W�v�ݒ�t�@�C���m" & tabinst_fn & "�n�ŁA�W�v����f�[�^�t�@�C�������͂���Ă��܂���B", vbExclamation, "MCS 2020 - Summarydata_Creation"
                        Application.StatusBar = False
                        End
                    End If
                    outdata_fn = ws_tabinst.Cells(2, 4)
                    weight_flg = ws_tabinst.Cells(2, 12)    ' �E�G�C�g�W�v�̗L���t���O�擾
                    If weight_flg = "" Then weight_flg = "�Ȃ�"
                Else
                    MsgBox "�W�v�ݒ�t�@�C���m" & tabinst_fn & "�n�����݂��܂���B", vbExclamation, "MCS 2020 - Summarydata_Creation"
                    Application.StatusBar = False
                    End
                End If
                
                Call Outdata_Open
                Call Setup_Hold
                
                wb_outdata.Activate
                ws_outdata.Select
                outdata_row = Cells(Rows.Count, setup_col).End(xlUp).Row
                outdata_col = Cells(1, Columns.Count).End(xlToLeft).Column
                
                If weight_flg = "����" Then   ' �E�G�C�g�i�␳�l�j�̎擾
                  w_index = Qcode_Match("weight")
                End If
                
                ' 2018/05/28 - �T�}���[�̏o�͐�́mSUM�n�t�H���_�Œ�ɕύX�B
'                summary_fd = Replace(Left(tabinst_fn, InStr(tabinst_fn, ".") - 1), ws_mainmenu.Cells(gcode_row, gcode_col), "")
                summary_fd = "SUM"
                summary_fn = Left(tabinst_fn, InStr(tabinst_fn, ".") - 1) & "_sum.xlsx"
                
'2018/04/26 - �ǋL ==========================
                Open file_path & "\" & summary_fd & "\" & summary_fn For Append As #1
                Close #1
                If Err.Number > 0 Then
                    Workbooks(summary_fn).Close
                End If
                If Dir(file_path & "\" & summary_fd & "\" & summary_fn) <> "" Then
                    Kill file_path & "\" & summary_fd & "\" & summary_fn
                End If
'============================================
                
                Call Cross_Setting(crs_cnt, fn_cnt)
                
                ' �f�[�^�t�@�C���̃N���[�Y
                wb_outdata.Activate
                ws_outdata.Select
                Columns("XFA:XFD").Select
                Selection.ClearContents
                Range("B7").Select
                Application.DisplayAlerts = False
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                Set wb_outdata = Nothing
                Set ws_outdata = Nothing
                
                ' �W�v�ݒ�t�@�C���̃N���[�Y
                wb_tabinst.Activate
                Application.DisplayAlerts = False
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                Set wb_tabinst = Nothing
                Set ws_tabinst = Nothing
                
'2020/01/10 - MCODE�����ǋL =================
                Call Mcode_Setting
'============================================
    
                ' �W�v�T�}���[�t�@�C����ۑ����ăN���[�Y
                wb_summary.Activate
                ws_summary1.Select
                
                If Dir(file_path & "\" & summary_fd, vbDirectory) = "" Then
                    MkDir file_path & "\" & summary_fd
                End If
                
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=file_path & "\" & summary_fd & "\" & summary_fn
                ActiveWorkbook.Close
                Application.DisplayAlerts = True
                
                crs_cnt = crs_cnt + 1
            Next i_cnt
            
            Set wb_summary = Nothing
            Set ws_summary0 = Nothing
            Set ws_summary1 = Nothing
            Set ws_summary2 = Nothing
            Set ws_summary3 = Nothing
            
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
              ws_mainmenu.Cells(41, 6) = "23"
            Else
              ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 23"
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
            Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�T�}���[�f�[�^�̍쐬�F�Ώۃt�@�C���mSUM�t�H���_����" & crs_cnt - 1 & "�̏W�v�ݒ�t�@�C���n"
            Close #1
            Call Finishing_Mcs2017
            MsgBox crs_cnt - 1 & "�̏W�v�T�}���[�f�[�^���������܂����B", vbInformation, "MCS 2020 - Summarydata_Creation"
            End
        ElseIf r_code = vbCancel Then
            Call Finishing_Mcs2017
            End
        End If
    End If
    
' �W�v�T�}���[�f�[�^�P��쐬����
    fn_cnt = 1
    crs_cnt = 1
    tabinst_fn = Application.GetOpenFilename("�W�v�ݒ�t�@�C��,*.xlsx", , "�W�v�ݒ�t�@�C�����J��")
    If tabinst_fn = "False" Then
        ' �L�����Z���{�^���̏���
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf tabinst_fn = "" Then
        MsgBox "�W�v����m�W�v�ݒ�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        wb.Activate
        ws_mainmenu.Select
        End
    End If
    
    ' �t���p�X����t�@�C�����̎擾
    tabinst_fn = Dir(tabinst_fn)

    wb.Activate
    ws_mainmenu.Select
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & tabinst_fn) <> "" Then
        Workbooks.Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & _
         ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & tabinst_fn
        Set wb_tabinst = Workbooks(tabinst_fn)
        Set ws_tabinst = wb_tabinst.ActiveSheet
        If ws_tabinst.Cells(2, 4) = "" Then
            MsgBox "�W�v�ݒ�t�@�C���m" & tabinst_fn & "�n�ŁA�W�v����f�[�^�t�@�C�������͂���Ă��܂���B", vbExclamation, "MCS 2020 - Summarydata_Creation"
            Application.StatusBar = False
            End
        End If
        outdata_fn = ws_tabinst.Cells(2, 4)
        weight_flg = ws_tabinst.Cells(2, 12)    ' �E�G�C�g�W�v�̗L���t���O�擾
        If weight_flg = "" Then weight_flg = "�Ȃ�"
    Else
        MsgBox "�W�v�ݒ�t�@�C���m" & tabinst_fn & "�n�����݂��܂���B", vbExclamation, "MCS 2020 - Summarydata_Creation"
        Application.StatusBar = False
        End
    End If
    
    Call Outdata_Open
    Call Setup_Hold
    
    wb_outdata.Activate
    ws_outdata.Select
    outdata_row = Cells(Rows.Count, setup_col).End(xlUp).Row
    outdata_col = Cells(1, Columns.Count).End(xlToLeft).Column
    
    If weight_flg = "����" Then   ' �E�G�C�g�i�␳�l�j�̎擾
        w_index = Qcode_Match("weight")
    End If
    
    ' 2018/05/28 - �T�}���[�̏o�͐�́mSUM�n�t�H���_�Œ�ɕύX�B
'    summary_fd = Replace(Left(tabinst_fn, InStr(tabinst_fn, ".") - 1), ws_mainmenu.Cells(gcode_row, gcode_col), "")
    summary_fd = "SUM"
    summary_fn = Left(tabinst_fn, InStr(tabinst_fn, ".") - 1) & "_sum.xlsx"
    
'2018/04/26 - �ǋL ==========================
    Open file_path & "\" & summary_fd & "\" & summary_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(summary_fn).Close
    End If
    If Dir(file_path & "\" & summary_fd & "\" & summary_fn) <> "" Then
        Kill file_path & "\" & summary_fd & "\" & summary_fn
    End If
'============================================
    
    Call Cross_Setting(crs_cnt, fn_cnt)
    
    ' �f�[�^�t�@�C���̃N���[�Y
    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFA:XFD").Select
    Selection.ClearContents
    Range("B7").Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
    
    ' �W�v�ݒ�t�@�C���̃N���[�Y
    wb_tabinst.Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    Set wb_tabinst = Nothing
    Set ws_tabinst = Nothing
    
'2020/01/10 - MCODE�����ǋL =================
    Call Mcode_Setting
'============================================
    
    ' �W�v�T�}���[�t�@�C����ۑ����ăN���[�Y
    wb_summary.Activate
    ws_summary1.Select
    
    If Dir(file_path & "\" & summary_fd, vbDirectory) = "" Then
        MkDir file_path & "\" & summary_fd
    End If
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=file_path & "\" & summary_fd & "\" & summary_fn
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Set wb_summary = Nothing
    Set ws_summary0 = Nothing
    Set ws_summary1 = Nothing
    Set ws_summary2 = Nothing
    Set ws_summary3 = Nothing
    
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
      ws_mainmenu.Cells(41, 6) = "23"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 23"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�T�}���[�f�[�^�̍쐬�F�Ώۃt�@�C���m" & tabinst_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�W�v�T�}���[�f�[�^���������܂����B", vbInformation, "MCS 2020 - Summarydata_Creation"
End Sub

Private Sub Cross_Setting(ByVal crs_cntx As Long, ByVal fn_cntx As Long)
' �W�v�T�}���[�t�@�C���̍쐬
    Dim waitTime As Variant
    Dim i_cnt As Long
    Dim f_cnt As Long
    Dim max_row As Long
    Dim select_state As String
    
    c_cnt = 1
    sum_row = 1: sum_col = 1
    div_row = 1: div_col = 1
     
    ' �W�v�T�}���[�o�͗p�t�@�C���̓W�J
    Workbooks.Add
    Worksheets.Add after:=Worksheets(1)
    Worksheets.Add after:=Worksheets(2)
    Worksheets.Add after:=Worksheets(3)
    
    Set wb_summary = ActiveWorkbook
    Set ws_summary0 = wb_summary.Worksheets("Sheet1")
    Set ws_summary1 = wb_summary.Worksheets("Sheet2")
    Set ws_summary2 = wb_summary.Worksheets("Sheet3")
    Set ws_summary3 = wb_summary.Worksheets("Sheet4")
    
    ws_summary0.Name = "�ڎ�"
    ws_summary1.Name = "�m���\"
    ws_summary2.Name = "�m�\"
    ws_summary3.Name = "���\"
    
    ' �W�v�ݒ�t�@�C���̏������[�U�[�^�ϐ��Ɋi�[
    wb_tabinst.Activate
    ws_tabinst.Select
    
    max_row = ws_tabinst.Cells(Rows.Count, setup_col).End(xlUp).Row
    ReDim c_data(max_row)
    
'2018/05/01 - �ǋL ==========================
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range(Cells(7, 1), Cells(max_row, 1)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range(Cells(6, 1), Cells(max_row, 21))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'============================================
    
    Application.ScreenUpdating = False
    Load UserForm52
    UserForm52.StartUpPosition = 1
    UserForm52.Show vbModeless
    UserForm52.Repaint
    Application.Visible = False
    AppActivate UserForm52.Caption
    UserForm52.Label2.Caption = tabinst_fn
    UserForm52.Label3.Caption = outdata_fn
    UserForm52.Label4.Caption = summary_fn
    UserForm52.Label5.Caption = "�o�͐�F" & file_path & "\" & summary_fd
    UserForm52.Label7.Caption = "[" & crs_cntx & "/" & fn_cntx & "�t�@�C��]"
    
    For i_cnt = 7 To max_row
        DoEvents
        UserForm52.Label1.Caption = Int(i_cnt / max_row * 100) & "%"
        UserForm52.Label6.Caption = "�W�v��" & Status_Dot(c_cnt)
        wb_tabinst.Activate
        ws_tabinst.Select
        If Mid(Cells(i_cnt, 1), 1, 1) <> "*" Then
            c_data(c_cnt).hyo_num = Cells(i_cnt, 1)
            c_data(c_cnt).f_code = Cells(i_cnt, 2)
            c_data(c_cnt).a_code = Cells(i_cnt, 3)
            c_data(c_cnt).r_code = Cells(i_cnt, 4)
            c_data(c_cnt).fna_flg = Cells(i_cnt, 5)
            c_data(c_cnt).ana_flg = Cells(i_cnt, 6)
            c_data(c_cnt).bosu_flg = Cells(i_cnt, 7)
            c_data(c_cnt).sum_flg = Cells(i_cnt, 8)
            c_data(c_cnt).ave_flg = Cells(i_cnt, 9)
            c_data(c_cnt).sd_flg = Cells(i_cnt, 10)
            c_data(c_cnt).min_flg = Cells(i_cnt, 11)
            c_data(c_cnt).q1_flg = Cells(i_cnt, 12)
            c_data(c_cnt).med_flg = Cells(i_cnt, 13)
            c_data(c_cnt).q3_flg = Cells(i_cnt, 14)
            c_data(c_cnt).max_flg = Cells(i_cnt, 15)
            c_data(c_cnt).mod_flg = Cells(i_cnt, 16)
            c_data(c_cnt).sel_code = Cells(i_cnt, 17)
            c_data(c_cnt).sel_value = Val(Cells(i_cnt, 18))
            c_data(c_cnt).ken_flg = Cells(i_cnt, 19)
            c_data(c_cnt).yuko_flg = Cells(i_cnt, 20)
            c_data(c_cnt).nobe_flg = Cells(i_cnt, 21)

            c_data(c_cnt).top1_flg = Cells(i_cnt, 22)
            c_data(c_cnt).sort_flg = Cells(i_cnt, 23)
            c_data(c_cnt).exct_flg = Cells(i_cnt, 24)
            c_data(c_cnt).graph_flg = Cells(i_cnt, 25)

            select_flg = 0
            select_state = "0000000"
            se_msg = ""
            fs1_msg = "": fs2_msg = "": fs3_msg = ""
            as1_msg = "": as2_msg = "": as3_msg = ""
            
' �W�v�ݒ�̏W�v�����i�Z���N�g�j���擾
            se_index = 0
            
            ' QCODE���T�[�`
            If c_data(c_cnt).sel_code <> "" Then
                select_flg = 1
                Mid(select_state, 1, 1) = "1"
                se_index = Qcode_Match(c_data(c_cnt).sel_code)
                
                ' �}���`�A���T�[�̏���
                se_cnt = 0
                If (q_data(se_index).q_format = "M") Or (Mid(q_data(se_index).q_format, 1, 1) = "L") Then
                    se_cnt = c_data(c_cnt).sel_value
                End If
                
                ' �\��ƃJ�e�S���[�̃R�����g�̐ݒ肪�Ȃ�������A�Z���N�g�R�����g�͔�\���ɂ���B
                If (q_data(se_index).q_title = "") And (q_data(se_index).q_ct(c_data(c_cnt).sel_value) = "") Then
                    se_msg = ""
                Else
                    se_msg = q_data(se_index).q_title & "�F" & q_data(se_index).q_ct(c_data(c_cnt).sel_value)
                End If
            End If

' �\���̐ݒ�����擾
            f_index = 0
            fs1_index = 0: fs2_index = 0: fs3_index = 0
            
            ' QCODE���T�[�`
            If c_data(c_cnt).f_code <> "" Then
                f_index = Qcode_Match(c_data(c_cnt).f_code)
                
                ' �}���`�A���T�[�̏���
                fma_cnt = 0
                If (q_data(f_index).q_format = "M") Or (Mid(q_data(se_index).q_format, 1, 1) = "L") Then
                    fma_cnt = q_data(f_index).ct_count
                End If
                
                ' �Z���N�g�@���T�[�`
                If q_data(f_index).sel_code1 <> "" Then
                    select_flg = 1
                    Mid(select_state, 2, 1) = "1"
                    fs1_index = Qcode_Match(q_data(f_index).sel_code1)
                    
                    ' �}���`�A���T�[�̏���
                    fs1_cnt = 0
                    If (q_data(fs1_index).q_format = "M") Or (Mid(q_data(fs1_index).q_format, 1, 1) = "L") Then
                        fs1_cnt = q_data(f_index).sel_value1
                    End If
                    fs1_msg = q_data(fs1_index).q_title & "�F" & q_data(fs1_index).q_ct(q_data(f_index).sel_value1)
                End If
                
                ' �Z���N�g�A���T�[�`
                If q_data(f_index).sel_code2 <> "" Then
                    select_flg = 1
                    Mid(select_state, 3, 1) = "1"
                    fs2_index = Qcode_Match(q_data(f_index).sel_code2)
                    
                    ' �}���`�A���T�[�̏���
                    fs2_cnt = 0
                    If (q_data(fs2_index).q_format = "M") Or (Mid(q_data(fs2_index).q_format, 1, 1) = "L") Then
                        fs2_cnt = q_data(f_index).sel_value2
                    End If
                    fs2_msg = q_data(fs2_index).q_title & "�F" & q_data(fs2_index).q_ct(q_data(f_index).sel_value2)
                End If
                
                ' �Z���N�g�B���T�[�`
                If q_data(f_index).sel_code3 <> "" Then
                    select_flg = 1
                    Mid(select_state, 4, 1) = "1"
                    fs3_index = Qcode_Match(q_data(f_index).sel_code3)
                    
                    ' �}���`�A���T�[�̏���
                    fs3_cnt = 0
                    If (q_data(fs3_index).q_format = "M") Or (Mid(q_data(fs3_index).q_format, 1, 1) = "L") Then
                        fs3_cnt = q_data(f_index).sel_value3
                    End If
                    fs3_msg = q_data(fs3_index).q_title & "�F" & q_data(fs3_index).q_ct(q_data(f_index).sel_value3)
                End If
            End If

' �\���̐ݒ�����擾
            a_index = 0
            ra_index = 0
            as1_index = 0: as2_index = 0: as3_index = 0
                
            If c_data(c_cnt).a_code <> "" Then
                a_index = Qcode_Match(c_data(c_cnt).a_code)
                
                ' ����QCODE���T�[�`
                If c_data(c_cnt).r_code <> "" Then
                    ra_index = Qcode_Match(c_data(c_cnt).r_code)
                
                    ' �������ڂƕ\�����ڂ̃Z���N�g�̃`�F�b�N
                    If q_data(a_index).sel_code1 <> q_data(ra_index).sel_code1 Then
                        MsgBox "�\�� QCODE�m" & q_data(a_index).q_code & "�n��" & vbCrLf & _
                        "���� QCODE�m" & q_data(ra_index).q_code & "�n��" & vbCrLf & _
                        "�����@�𓯂��ݒ�ɂ��Ă��������B" & vbCrLf & vbCrLf & _
                        "�yTIPS�z" & vbCrLf & "�ݒ��ʂŁA��L QCODE �̏����@�̐ݒ���m�F���Ă��������B", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                    If q_data(a_index).sel_code2 <> q_data(ra_index).sel_code2 Then
                        MsgBox "�\�� QCODE�m" & q_data(a_index).q_code & "�n��" & vbCrLf & _
                        "���� QCODE�m" & q_data(ra_index).q_code & "�n��" & vbCrLf & _
                        "�����A�𓯂��ݒ�ɂ��Ă��������B" & vbCrLf & vbCrLf & _
                        "�yTIPS�z" & vbCrLf & "�ݒ��ʂŁA��L QCODE �̏����A�̐ݒ���m�F���Ă��������B", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                    If q_data(a_index).sel_code3 <> q_data(ra_index).sel_code3 Then
                        MsgBox "�\�� QCODE�m" & q_data(a_index).q_code & "�n��" & vbCrLf & _
                        "���� QCODE�m" & q_data(ra_index).q_code & "�n��" & vbCrLf & _
                        "�����B�𓯂��ݒ�ɂ��Ă��������B" & vbCrLf & vbCrLf & _
                        "�yTIPS�z" & vbCrLf & "�ݒ��ʂŁA��L QCODE �̏����B�̐ݒ���m�F���Ă��������B", vbExclamation, "MCS 2020 - Cross_Setting"
                        Call Files_Close
                        End
                    End If
                End If
                
                ' �}���`�A���T�[�̏���
                ama_cnt = 0
                If (q_data(a_index).q_format = "M") Or (Mid(q_data(a_index).q_format, 1, 1) = "L") Then
                    ama_cnt = q_data(a_index).ct_count
                End If
                
                ' �Z���N�g�@���T�[�`
                If q_data(a_index).sel_code1 <> "" Then
                    select_flg = 1
                    Mid(select_state, 5, 1) = "1"
                    as1_index = Qcode_Match(q_data(a_index).sel_code1)
                    
                    ' �}���`�A���T�[�̏���
                    as1_cnt = 0
                    If (q_data(as1_index).q_format = "M") Or (Mid(q_data(as1_index).q_format, 1, 1) = "L") Then
                        as1_cnt = q_data(a_index).sel_value1
                    End If
                    as1_msg = q_data(as1_index).q_title & "�F" & q_data(as1_index).q_ct(q_data(a_index).sel_value1)
                End If
                
                ' �Z���N�g�A���T�[�`
                If q_data(a_index).sel_code2 <> "" Then
                    select_flg = 1
                    Mid(select_state, 6, 1) = "1"
                    as2_index = Qcode_Match(q_data(a_index).sel_code2)
                    
                    ' �}���`�A���T�[�̏���
                    as2_cnt = 0
                    If (q_data(as2_index).q_format = "M") Or (Mid(q_data(as2_index).q_format, 1, 1) = "L") Then
                        as2_cnt = q_data(a_index).sel_value2
                    End If
                    as2_msg = q_data(as2_index).q_title & "�F" & q_data(as2_index).q_ct(q_data(a_index).sel_value2)
                End If
                
                ' �Z���N�g�B���T�[�`
                If q_data(a_index).sel_code3 <> "" Then
                    select_flg = 1
                    Mid(select_state, 7, 1) = "1"
                    as3_index = Qcode_Match(q_data(a_index).sel_code3)
                    
                    ' �}���`�A���T�[�̏���
                    as3_cnt = 0
                    If (q_data(as3_index).q_format = "M") Or (Mid(q_data(as3_index).q_format, 1, 1) = "L") Then
                        as3_cnt = q_data(a_index).sel_value3
                    End If
                    as3_msg = q_data(as3_index).q_title & "�F" & q_data(as3_index).q_ct(q_data(a_index).sel_value3)
                End If
            Else
                MsgBox "�\���́mQCODE�n�́A�K���w�肵�Ă��������B", vbExclamation, "MCS 2020 - Cross_Setting"
                Call Files_Close
                End
            End If
            
            wb_summary.Activate
            Call Cross_Index                        ' �ڎ�������
            Call Cross_Header                       ' �W�v�\�w�b�_�[�R�����g������
            
            If select_flg = 1 Then
                Call Select_Flag(select_state)      ' �Z���N�g������
            End If
            
            If weight_flg = "����" Then
                Call weight_ra                      ' �E�G�C�g�W�v�E�����l�Z�o������
            End If
            
            If c_data(c_cnt).fna_flg <> "E" Then
                Call Simple_Summary                 ' �W�v�l�Z�o�E�����ݖ⏈���ցi�P���W�v�j
            Else
                sum_row = sum_row - 2
                div_row = div_row - 1
            End If
            
            If c_data(c_cnt).f_code <> "" Then
                For f_cnt = 1 To q_data(f_index).ct_count
                    sum_row = sum_row + 2
                    div_row = div_row + 1
                    Call Cross_Tabulation(f_cnt)    ' �W�v�l�Z�o�E�����ݖ⏈���ցi�N���X�W�v�j
                Next f_cnt
                
                If c_data(c_cnt).fna_flg <> "N" Then
                    sum_row = sum_row + 2
                    div_row = div_row + 1
                    Call FaceNa_Tabulation          ' �W�v�l�Z�o�E�����ݖ⏈���ցi�N���X�W�v�j���\������
                End If
            End If
            
            sum_row = sum_row + 3
            sum_col = 1
            
            div_row = div_row + 2
            div_col = 1
            
            c_cnt = c_cnt + 1
        End If
    Next i_cnt
    UserForm52.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")
    Application.Wait waitTime
    Application.Visible = True
    Unload UserForm52
End Sub

Private Sub Files_Close()
' �G���[�I�����̃t�@�C���N���[�Y
    Application.DisplayAlerts = False
    
    ' �f�[�^�t�@�C��
    wb_outdata.Activate
    ActiveWorkbook.Close
    
    ' �W�v�ݒ�t�@�C��
    wb_tabinst.Activate
    ActiveWorkbook.Close
    
    ' �W�v�T�}���[�t�@�C��
    wb_summary.Activate
    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True
    
    wb.Activate
    ws_setup.Select
    ws_setup.Cells(1, 1).Select
    ws_mainmenu.Select
    ws_mainmenu.Cells(3, 8).Select
    Application.StatusBar = False
    Application.Visible = True
    Unload UserForm52
End Sub

Private Sub Cross_Index()
    Dim sel_com As String
' �ڎ��̍쐬
    If c_cnt = 1 Then
        ws_summary0.Cells(1, 1) = "�A��"
        ws_summary0.Cells(1, 2) = "�\��"
        ws_summary0.Cells(1, 3) = "MCODE"
        ws_summary0.Cells(1, 4) = "�\���i�c���j"
        ws_summary0.Cells(1, 5) = "�\���i�����j"
        ws_summary0.Cells(1, 6) = "�W�v����"
        ws_summary0.Cells(1, 7) = "�����N"
    
        If weight_flg = "����" Then   ' �E�G�C�g�i�␳�l�j�̎擾
            ws_summary0.Cells(1, 1) = "�A�ԃE��"
        End If
    End If
    
    ws_summary0.Cells(c_cnt + 1, 1) = c_cnt
    ws_summary0.Cells(c_cnt + 1, 2) = "'" & c_data(c_cnt).hyo_num
    ws_summary0.Cells(c_cnt + 1, 3) = q_data(a_index).m_code
    ws_summary0.Cells(c_cnt + 1, 4) = q_data(f_index).q_title
    ws_summary0.Cells(c_cnt + 1, 5) = q_data(a_index).q_title
    
    sel_com = ""
    If se_msg <> "" Then
        sel_com = sel_com & se_msg
    End If
    
    If fs1_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs1_msg
        Else
            sel_com = sel_com & vbLf & fs1_msg
        End If
    End If

    If fs2_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs2_msg
        Else
            sel_com = sel_com & vbLf & fs2_msg
        End If
    End If

    If fs3_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & fs3_msg
        Else
            sel_com = sel_com & vbLf & fs3_msg
        End If
    End If

    If as1_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as1_msg
        Else
            sel_com = sel_com & vbLf & as1_msg
        End If
    End If

    If as2_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as2_msg
        Else
            sel_com = sel_com & vbLf & as2_msg
        End If
    End If

    If as3_msg <> "" Then
        If sel_com = "" Then
            sel_com = sel_com & as3_msg
        Else
            sel_com = sel_com & vbLf & as3_msg
        End If
    End If
    ws_summary0.Cells(c_cnt + 1, 6) = sel_com

    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 7), _
     Address:="", SubAddress:="'" & ws_summary1.Name & "'!A" & sum_row, TextToDisplay:="�m���\"
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 8), _
     Address:="", SubAddress:="'" & ws_summary2.Name & "'!A" & div_row, TextToDisplay:="�m�\"
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary0.Cells(c_cnt + 1, 9), _
     Address:="", SubAddress:="'" & ws_summary3.Name & "'!A" & div_row, TextToDisplay:="���\"
End Sub

Private Sub Cross_Header()
    Dim i_cnt As Long
    Dim temp_row As Long
    Dim temp_col As Long
    Dim unit_cm As String
    Dim format_cm As String
    Dim bosu_cm As String
' �W�v�\�w�b�_�[�R�����g����
    ' �\���̏���
    ' ���L�̓n�C�p�[�����N������Ƃ��ďo�͂���p�^�[��
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary1.Cells(sum_row, sum_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!G" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary2.Cells(div_row, div_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!H" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    
    ActiveSheet.Hyperlinks.Add Anchor:=ws_summary3.Cells(div_row, div_col), _
     Address:="", SubAddress:="'" & ws_summary0.Name & "'!I" & c_cnt + 1, TextToDisplay:="'" & c_data(c_cnt).hyo_num
    ' ���L�͂��̂܂܃e�L�X�g�𕶎���Ƃ��ďo�͂���p�^�[��
    'ws_summary1.Cells(sum_row, sum_col).Value = "'" & c_data(c_cnt).hyo_num
    'ws_summary2.Cells(div_row, div_col).Value = "'" & c_data(c_cnt).hyo_num
    'ws_summary3.Cells(div_row, div_col).Value = "'" & c_data(c_cnt).hyo_num
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' MCODE�̏���
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).m_code
    ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).m_code
    ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).m_code
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    
    ' �\��̏���
    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).q_title
    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).q_title
    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).q_title
    sum_row = sum_row + 1
    div_row = div_row + 1
    
' �Z���N�g�R�����g�̏o�͏���
' �y�P�Ԗځz�W�v�ݒ�Z���N�g
' �y�Q�Ԗځz�\���Z���N�g�@
' �y�R�Ԗځz�\���Z���N�g�A
' �y�S�Ԗځz�\���Z���N�g�B
' �y�T�Ԗځz�\���Z���N�g�@
' �y�U�Ԗځz�\���Z���N�g�A
' �y�V�Ԗځz�\���Z���N�g�B
'
    ' �W�v�ݒ�Z���N�g�R�����g�̏���
    If se_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�W�v�����z" & se_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�W�v�����z" & se_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�W�v�����z" & se_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�@�R�����g�̏���
    If fs1_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & fs1_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs1_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs1_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�A�R�����g�̏���
    If fs2_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & fs2_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs2_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs2_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�B�R�����g�̏���
    If fs3_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & fs3_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs3_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & fs3_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�@�R�����g�̏���
    If as1_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & as1_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as1_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as1_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�A�R�����g�̏���
    If as2_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & as2_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as2_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as2_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If
    
    ' �\���Z���N�g�B�R�����g�̏���
    If as3_msg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = "�y�\���W�v�����z" & as3_msg
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as3_msg
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = "�y�\���W�v�����z" & as3_msg
        sum_row = sum_row + 1
        div_row = div_row + 1
    End If

'---------------------------------
' �g���I�v�V�����̏��� - 2018.9.14
    ' TOP1�E�}�[�L���O�t���O
    ws_summary1.Cells(sum_row, sum_col - 2).Value = c_data(c_cnt).top1_flg
    ws_summary2.Cells(div_row, div_col - 2).Value = c_data(c_cnt).top1_flg
    ws_summary3.Cells(div_row, div_col - 2).Value = c_data(c_cnt).top1_flg

    ' CT�\�[�g�E�~���t���O�����OCT�t���O
    ws_summary1.Cells(sum_row + 1, sum_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg
    ws_summary2.Cells(div_row + 1, div_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg
    ws_summary3.Cells(div_row + 1, div_col - 2).Value = c_data(c_cnt).sort_flg & c_data(c_cnt).exct_flg

    ' �O���t�E�쐬�t���O
    ws_summary1.Cells(sum_row + 2, sum_col - 2).Value = c_data(c_cnt).graph_flg
    ws_summary2.Cells(div_row + 2, div_col - 2).Value = c_data(c_cnt).graph_flg
    ws_summary3.Cells(div_row + 2, div_col - 2).Value = c_data(c_cnt).graph_flg
'---------------------------------
    
    ' �\���R�����g�̏���
    temp_row = sum_row
    temp_col = sum_col
    
    ' �\���J�e�S���[�ԍ��̏���
    sum_col = sum_col + 3
    div_col = div_col + 3
    For i_cnt = 1 To q_data(a_index).ct_count
        ws_summary1.Cells(sum_row, sum_col).Value = i_cnt
        ws_summary2.Cells(div_row, div_col).Value = i_cnt
        ws_summary3.Cells(div_row, div_col).Value = i_cnt
        sum_col = sum_col + 1
        div_col = div_col + 1
    Next i_cnt
    
    ' �J�e�S���[�ԍ��m���񓚁n�̏���
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "N/A"
        ws_summary2.Cells(div_row, div_col).Value = "N/A"
        ws_summary3.Cells(div_row, div_col).Value = "N/A"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    sum_row = sum_row + 1
    sum_col = temp_col
    div_row = div_row + 1
    div_col = temp_col
    
    ' �R�����g�m�ݖ�`���n�Ɓm�\����ꐔ�n�̏��� - 2018.5.24 �ǉ�
    format_cm = ""
    If q_data(a_index).q_format = "S" Then
        format_cm = "�m�ݖ�`���F�P��񓚁n"
    ElseIf q_data(a_index).q_format = "M" Then
        format_cm = "�m�ݖ�`���F�����񓚁n"
    ElseIf Mid(q_data(a_index).q_format, 1, 1) = "L" Then
        If q_data(a_index).ct_loop = 1 Then
            format_cm = "�m�ݖ�`���F�P��񓚁n"
        Else
            format_cm = "�m�ݖ�`���F���蕡���񓚁n"
        End If
    ElseIf q_data(a_index).q_format = "R" Then
        format_cm = "�m�ݖ�`���F�����񓚁n"
    ElseIf q_data(a_index).q_format = "H" Then
        format_cm = "�m�ݖ�`���F�g�J�[�\���n"
    End If
    
    bosu_cm = ""
    If c_data(c_cnt).bosu_flg = "Y" Then
        bosu_cm = "�m�\����ꐔ�F�L���񓚐��n"
    Else
        bosu_cm = "�m�\����ꐔ�F�S�́n"
    End If
    
    ws_summary1.Cells(sum_row, sum_col).Value = format_cm & vbCrLf & bosu_cm
    ws_summary2.Cells(div_row, div_col).Value = format_cm & vbCrLf & bosu_cm
    ws_summary3.Cells(div_row, div_col).Value = format_cm & vbCrLf & bosu_cm
    
    ' �R�����g�m�S�́n�̏���
    ws_summary1.Cells(sum_row, sum_col + 2).Value = "����"
    ws_summary2.Cells(div_row, div_col + 2).Value = "����"
    ws_summary3.Cells(div_row, div_col + 2).Value = "����"
    sum_col = sum_col + 3
    div_col = div_col + 3
    
    ' �R�����g�m�\���n�̏���
    For i_cnt = 1 To q_data(a_index).ct_count
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = q_data(a_index).q_ct(i_cnt)
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = q_data(a_index).q_ct(i_cnt)
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = q_data(a_index).q_ct(i_cnt)
        sum_col = sum_col + 1
        div_col = div_col + 1
    Next i_cnt
    
    ' �R�����g�m���񓚁n�̏���
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "����"
        ws_summary2.Cells(div_row, div_col).Value = "����"
        ws_summary3.Cells(div_row, div_col).Value = "����"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�L���񓚗��n�̏���
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�L���񓚐�"
        ws_summary2.Cells(div_row, div_col).Value = "�L���񓚐�"
        ws_summary3.Cells(div_row, div_col).Value = "�L���񓚐�"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m���׉񓚗��n�̏���
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "���׉񓚐�"
        ws_summary2.Cells(div_row, div_col).Value = "���׉񓚐�"
        ws_summary3.Cells(div_row, div_col).Value = "���׉񓚐�"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m���v�n�̏���
    If c_data(c_cnt).sum_flg = "Y" Then
        If q_data(ra_index).r_unit <> "" Then
            unit_cm = "�i" & q_data(ra_index).r_unit & "�j"
        Else
            unit_cm = ""
        End If
        ws_summary1.Cells(sum_row, sum_col).Value = "���v" & unit_cm
        ws_summary2.Cells(div_row, div_col).Value = "���v" & unit_cm
        ws_summary3.Cells(div_row, div_col).Value = "���v" & unit_cm
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m���ρn�̏���
    If c_data(c_cnt).ave_flg <> "" Then
        If q_data(ra_index).r_unit <> "" Then
            unit_cm = "�i" & q_data(ra_index).r_unit & "�j"
        Else
            unit_cm = ""
        End If
        ws_summary1.Cells(sum_row, sum_col).Value = "����" & unit_cm
        ws_summary2.Cells(div_row, div_col).Value = "����" & unit_cm
        ws_summary3.Cells(div_row, div_col).Value = "����" & unit_cm
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�W���΍��n�̏���
    If c_data(c_cnt).sd_flg <> "" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�W���΍�"
        ws_summary2.Cells(div_row, div_col).Value = "�W���΍�"
        ws_summary3.Cells(div_row, div_col).Value = "�W���΍�"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�ŏ��l�n�̏���
    If c_data(c_cnt).min_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�ŏ��l"
        ws_summary2.Cells(div_row, div_col).Value = "�ŏ��l"
        ws_summary3.Cells(div_row, div_col).Value = "�ŏ��l"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m��P�l���ʁn�̏���
    If c_data(c_cnt).q1_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "��P�l����"
        ws_summary2.Cells(div_row, div_col).Value = "��P�l����"
        ws_summary3.Cells(div_row, div_col).Value = "��P�l����"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�����l�n�̏���
    If c_data(c_cnt).med_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�����l"
        ws_summary2.Cells(div_row, div_col).Value = "�����l"
        ws_summary3.Cells(div_row, div_col).Value = "�����l"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m��R�l���ʁn�̏���
    If c_data(c_cnt).q1_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "��R�l����"
        ws_summary2.Cells(div_row, div_col).Value = "��R�l����"
        ws_summary3.Cells(div_row, div_col).Value = "��R�l����"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�ő�l�n�̏���
    If c_data(c_cnt).max_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�ő�l"
        ws_summary2.Cells(div_row, div_col).Value = "�ő�l"
        ws_summary3.Cells(div_row, div_col).Value = "�ő�l"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �R�����g�m�ŕp�l�n�̏���
    If c_data(c_cnt).mod_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = "�ŕp�l"
        ws_summary2.Cells(div_row, div_col).Value = "�ŕp�l"
        ws_summary3.Cells(div_row, div_col).Value = "�ŕp�l"
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    sum_row = sum_row + 1
    sum_col = temp_col
    div_row = div_row + 1
    div_col = temp_col
End Sub

Private Sub Select_Flag(ByVal state_flag As Long)
    Dim i_cnt As Long
' �W�v�Ώۃf�[�^�̂��Ƃ̃Z���N�g�t���O���H
' �y�T�v�z�W�v�Ώۃf�[�^�ƂȂ�t�@�C���̍ŏI��m16384�n��ڂɁA�Z���N�g�����𖞂����Ă���΃t���O�����Ă�B
'
    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFD:XFD").Select
    Selection.ClearContents
    Cells(6, 16384) = "Select"
    
    ' �T���v���x�[�X�ł̃Z���N�g�̗L������
    For i_cnt = 7 To outdata_row
        Select Case state_flag
        Case "0000000"
            '�����Ȃ��i�Z���N�g�Ȃ��j
        Case "1111111"
            Call Select_Decision01(i_cnt)
        Case "1111110"
            Call Select_Decision02(i_cnt)
        Case "1111100"
            Call Select_Decision03(i_cnt)
        Case "1111000"
            Call Select_Decision04(i_cnt)
        Case "1110000"
            Call Select_Decision05(i_cnt)
        Case "1100000"
            Call Select_Decision06(i_cnt)
        Case "1000000"
            Call Select_Decision07(i_cnt)
        Case "1110111"
            Call Select_Decision08(i_cnt)
        Case "1100111"
            Call Select_Decision09(i_cnt)
        Case "1000111"
            Call Select_Decision10(i_cnt)
        Case "1110110"
            Call Select_Decision11(i_cnt)
        Case "1110100"
            Call Select_Decision12(i_cnt)
        Case "1100110"
            Call Select_Decision13(i_cnt)
        Case "1100100"
            Call Select_Decision14(i_cnt)
        Case "1000110"
            Call Select_Decision15(i_cnt)
        Case "1000100"
            Call Select_Decision16(i_cnt)
        Case "0111111"
            Call Select_Decision17(i_cnt)
        Case "0111110"
            Call Select_Decision18(i_cnt)
        Case "0111100"
            Call Select_Decision19(i_cnt)
        Case "0111000"
            Call Select_Decision20(i_cnt)
        Case "0110000"
            Call Select_Decision21(i_cnt)
        Case "0100000"
            Call Select_Decision22(i_cnt)
        Case "0110111"
            Call Select_Decision23(i_cnt)
        Case "0100111"
            Call Select_Decision24(i_cnt)
        Case "0000111"
            Call Select_Decision25(i_cnt)
        Case "0110110"
            Call Select_Decision26(i_cnt)
        Case "0110100"
            Call Select_Decision27(i_cnt)
        Case "0100110"
            Call Select_Decision28(i_cnt)
        Case "0100100"
            Call Select_Decision29(i_cnt)
        Case "0000110"
            Call Select_Decision30(i_cnt)
        Case "0000100"
            Call Select_Decision31(i_cnt)
        End Select
    Next i_cnt
End Sub

Private Sub Select_Decision01(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���B�E�\���@�E�\���A�E�\���B
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    If as3_cnt = 0 Then
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    Else
                                                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision02(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���B�E�\���@�E�\���A
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            If as2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision03(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���B�E�\���@
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    If as1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision04(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���B
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If fs3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision05(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision06(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision07(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Select_Decision08(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���@�E�\���A�E�\���B
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            If fs2_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision09(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���@�E�\���A�E�\���B
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    If fs1_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision10(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���B
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If as3_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision11(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���@�E�\���A
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision12(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A�E�\���@
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If fs2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision13(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���@�E�\���A
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision14(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���@
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If fs1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision15(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@�E�\���A
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    If as2_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision16(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �W�v�ݒ�E�\���@
    If se_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column) = c_data(c_cnt).sel_value Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(se_index).data_column + se_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision17(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���B�E�\���@�E�\���A�E�\���B
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            If as3_cnt = 0 Then
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            Else
                                                If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision18(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���B�E�\���@�E�\���A
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    If as2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision19(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���B�E�\���@
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            If as1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision20(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���B
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If fs3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column) = q_data(f_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs3_index).data_column + fs3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision21(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision22(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Select_Decision23(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���@�E�\���A�E�\���B
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    If fs2_cnt = 0 Then
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    Else
                                        If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                                            ws_outdata.Cells(ix_cnt, 16384) = 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision24(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���@�E�\���A�E�\���B
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            If fs1_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision25(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���B
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If as3_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column) = q_data(a_index).sel_value3 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as3_index).data_column + as3_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision26(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���@�E�\���A
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            If as2_cnt = 0 Then
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            Else
                                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                                    ws_outdata.Cells(ix_cnt, 16384) = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision27(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A�E�\���@
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If fs2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column) = q_data(f_index).sel_value2 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(fs2_index).data_column + fs2_cnt - 1) = 1 Then
                    If as1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision28(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���@�E�\���A
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    If fs1_cnt = 0 Then
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    Else
                        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
                            ws_outdata.Cells(ix_cnt, 16384) = 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision29(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���@
    If fs1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column) = q_data(f_index).sel_value1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(fs1_index).data_column + fs1_cnt - 1) = 1 Then
            If as1_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision30(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@�E�\���A
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            If as2_cnt = 0 Then
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column) = q_data(a_index).sel_value2 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            Else
                If ws_outdata.Cells(ix_cnt, q_data(as2_index).data_column + as2_cnt - 1) = 1 Then
                    ws_outdata.Cells(ix_cnt, 16384) = 1
                End If
            End If
        End If
    End If
End Sub

Private Sub Select_Decision31(ByVal ix_cnt As Long)
' �Z���N�g����
' ���ꂼ��̃J�E���^�m*_cnt�n�́A�Y�����ڂ̂r�`�i*_cnt=0�j�E�l�`�i*_cnt<>0�j����ł��B
' �\���@
    If as1_cnt = 0 Then
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column) = q_data(a_index).sel_value1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    Else
        If ws_outdata.Cells(ix_cnt, q_data(as1_index).data_column + as1_cnt - 1) = 1 Then
            ws_outdata.Cells(ix_cnt, 16384) = 1
        End If
    End If
End Sub

Private Sub Simple_Summary()
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim gt_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
    Dim decimal_places As Integer
' �W�v�l�Z�o�E�����ݖ�̏����i�P���W�v�j
    On Error Resume Next
    temp_col = sum_col
    
    ' �m�S�́n�̎Z�o
    gt_cnt = 0
    If select_flg = 1 Then
        If weight_flg = "�Ȃ�" Then
            gt_cnt = Application.WorksheetFunction. _
             CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
             Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
        Else
            gt_cnt = Application.WorksheetFunction. _
             SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
             Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), "<>", _
             Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
        End If
    Else
        If weight_flg = "�Ȃ�" Then
            gt_cnt = Application.WorksheetFunction. _
             CountIf(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>" & "")
        Else
            gt_cnt = Application.WorksheetFunction. _
             SumIf(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), "<>", _
             Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
        End If
    End If
    
    ' �m���񓚁n�̎Z�o
    na_cnt = 0
    If ama_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "�Ȃ�" Then
                na_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                na_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "�Ȃ�" Then
                na_cnt = Application.WorksheetFunction. _
                 CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "")
            Else
                na_cnt = Application.WorksheetFunction. _
                 SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(a_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            na_cnt = na_cnt + 1
                        End If
                    End If
                Else
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(a_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                        End If
                    End If
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    For ma_cnt = 1 To q_data(a_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        na_cnt = na_cnt + 1
                    End If
                Else
                    For ma_cnt = 1 To q_data(a_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' �m�L���񓚐��n�̎Z�o
    vr_cnt = gt_cnt - na_cnt
    
    ' �����m�S�́n�̏���
    ws_summary1.Cells(sum_row, sum_col - 1).Value = "0"
    ws_summary2.Cells(div_row, div_col - 1).Value = "0"
    ws_summary3.Cells(div_row, div_col - 1).Value = "0"
    ws_summary1.Cells(sum_row, sum_col).Value = "�@�S�@��"
    ws_summary2.Cells(div_row, div_col).Value = "�@�S�@��"
    ws_summary3.Cells(div_row, div_col).Value = "�@�S�@��"
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' �������Ɂm�L���񓚐��n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' �������Ɂm�S�́n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = gt_cnt
        ws_summary2.Cells(div_row, div_col).Value = gt_cnt
        ws_summary3.Cells(div_row, div_col).Value = gt_cnt
    End If
    
    ' �E�G�C�g�W�v���m�S�́n�̃Z�������ݒ�
    If weight_flg = "����" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' �����m�J�e�S���[�n�̏���
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If gt_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1)
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     CountIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1)
                Else
                    ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                    ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                     SumIf(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)))
                    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If gt_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / gt_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' �����m���񓚁n�̏���
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' �E�G�C�g�W�v���m���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If gt_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / gt_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / gt_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m�L���񓚁n�̏���
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' �E�G�C�g�W�v���m�L���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If gt_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / gt_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / gt_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m���׉񓚁n�̏���
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' �E�G�C�g�W�v���m���׉񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If gt_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / gt_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / gt_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �W�v�ݒ�t�@�C���̏�������m�����ݖ�p�n
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If select_flg = 1 Then
        Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    End If
    
    If weight_flg = "�Ȃ�" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' �e�����ݖ�̏�����
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' �e�����ݖ�i�E�G�C�g�W�v�j�̏�����
    End If
    
    
    '�I�[�g�t�B���^�̉���
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub Cross_Tabulation(ByVal x_cnt As Long)
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim face_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
    Dim decimal_places As Integer
' �W�v�l�Z�o�E�����ݖ�̏����i�N���X�W�v�j
    On Error Resume Next
    temp_col = sum_col

    ' �m�\�����ځn�̎Z�o
    face_cnt = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            End If
        End If
    Else
        If select_flg = 1 Then
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            End If
        End If
    End If
    
    ' �m���񓚁n�̎Z�o
    na_cnt = 0
    If ama_cnt = 0 Then
        If fma_cnt = 0 Then
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                End If
            End If
        Else
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                End If
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If fma_cnt = 0 Then
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column) = x_cnt Then
                    If select_flg = 1 Then
                        If weight_flg = "�Ȃ�" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "�Ȃ�" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column + x_cnt - 1) = 1 Then
                    If select_flg = 1 Then
                        If weight_flg = "�Ȃ�" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "�Ȃ�" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' �m�L���񓚐��n�̎Z�o
    vr_cnt = face_cnt - na_cnt
    
    ' �\���w�b�_�[�R�����g�̏���
    If x_cnt = 1 Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
        ws_summary1.Cells(sum_row, sum_col).Value = q_data(f_index).q_title
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary2.Cells(div_row, div_col).Value = q_data(f_index).q_title
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
        ws_summary3.Cells(div_row, div_col).Value = q_data(f_index).q_title
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' �����m�\�����ځn�̏���
    ws_summary1.Cells(sum_row, sum_col - 2).Value = x_cnt
    ws_summary2.Cells(div_row, div_col - 2).Value = x_cnt
    ws_summary3.Cells(div_row, div_col - 2).Value = x_cnt
    ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "@"
    ws_summary1.Cells(sum_row, sum_col).Value = q_data(f_index).q_ct(x_cnt)
    ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary2.Cells(div_row, div_col).Value = q_data(f_index).q_ct(x_cnt)
    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "@"
    ws_summary3.Cells(div_row, div_col).Value = q_data(f_index).q_ct(x_cnt)
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' �������Ɂm�L���񓚐��n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' �������Ɂm�\�����ڑS���n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = face_cnt
        ws_summary2.Cells(div_row, div_col).Value = face_cnt
        ws_summary3.Cells(div_row, div_col).Value = face_cnt
    End If
    
    ' �E�G�C�g�W�v���m�\�����ڑS���n�̃Z�������ݒ�
    If weight_flg = "����" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' �����m�J�e�S���[�n�̏���
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), x_cnt)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column + x_cnt - 1), ws_outdata.Cells(outdata_row, q_data(f_index).data_column + x_cnt - 1)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' �����m���񓚁n�̏���
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' �E�G�C�g�W�v���m���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If face_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m�L���񓚁n�̏���
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' �E�G�C�g�W�v���m�L���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m���׉񓚁n�̏���
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' �E�G�C�g�W�v���m���׉񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �W�v�ݒ�t�@�C���̏�������m�����ݖ�p�n
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, x_cnt, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, x_cnt, visibledropdown:=False
        End If
    Else
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column + x_cnt - 1, 1, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column + x_cnt - 1, 1, visibledropdown:=False
        End If
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    End If
    
    If weight_flg = "�Ȃ�" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' �e�����ݖ�̏�����
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' �e�����ݖ�i�E�G�C�g�W�v�j�̏�����
    End If
    
    '�I�[�g�t�B���^�̉���
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub FaceNa_Tabulation()
    Dim a_cnt As Long
    Dim i_cnt As Long
    Dim face_cnt As Double
    Dim na_cnt As Double
    Dim vr_cnt As Double
    Dim total_cnt As Double
    Dim ma_cnt As Long
    Dim mna_cnt As Long
    Dim item_cnt As Double
    Dim item_per As Double
    Dim filter_flg As Integer
    Dim temp_col As Long
' �W�v�l�Z�o�E�����ݖ�̏����i�N���X�W�v�j���\������
' �y�T�v�z�t�@�C���́m16383�n��ڂɁA�\�����񓚃t���O�����Ă�B
'
    On Error Resume Next
    temp_col = sum_col

    wb_outdata.Activate
    ws_outdata.Select
    Columns("XFC:XFC").Select
    Selection.ClearContents
    Cells(6, 16383) = "FMA[N/A]"

    ' �m�\�����񓚍��ځn�̎Z�o
    face_cnt = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                 Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            End If
        Else
            If weight_flg = "�Ȃ�" Then
                face_cnt = Application.WorksheetFunction. _
                 CountIfs(Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            Else
                face_cnt = Application.WorksheetFunction. _
                 SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                 Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1)), "<>&""", _
                 Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(f_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            face_cnt = face_cnt + 1
                            ws_outdata.Cells(i_cnt, 16383) = 1
                        End If
                    End If
                Else
                    If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                        For ma_cnt = 1 To q_data(f_index).ct_count
                            mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                        Next ma_cnt
                        If mna_cnt = 0 Then
                            face_cnt = face_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            ws_outdata.Cells(i_cnt, 16383) = 1
                        End If
                    End If
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    For ma_cnt = 1 To q_data(f_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        face_cnt = face_cnt + 1
                        ws_outdata.Cells(i_cnt, 16383) = 1
                    End If
                Else
                    For ma_cnt = 1 To q_data(f_index).ct_count
                        mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(f_index).data_column + ma_cnt - 1))
                    Next ma_cnt
                    If mna_cnt = 0 Then
                        face_cnt = face_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                        ws_outdata.Cells(i_cnt, 16383) = 1
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' �m���񓚁n�̎Z�o
    na_cnt = 0
    If ama_cnt = 0 Then
        If fma_cnt = 0 Then
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                   na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                   na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                End If
            End If
        Else
            If select_flg = 1 Then
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                     Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                End If
            Else
                If weight_flg = "�Ȃ�" Then
                    na_cnt = Application.WorksheetFunction. _
                     CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                Else
                    na_cnt = Application.WorksheetFunction. _
                     SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                     Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), "", _
                     Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                End If
            End If
        End If
    Else
        For i_cnt = 7 To outdata_row
            mna_cnt = 0
            If fma_cnt = 0 Then
                If ws_outdata.Cells(i_cnt, q_data(f_index).data_column) = "" Then
                    If select_flg = 1 Then
                        If weight_flg = "�Ȃ�" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "�Ȃ�" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            Else
                If ws_outdata.Cells(i_cnt, 16383) = 1 Then
                    If select_flg = 1 Then
                        If weight_flg = "�Ȃ�" Then
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + 1
                                End If
                            End If
                        Else
                            If ws_outdata.Cells(i_cnt, 16384) = 1 Then
                                For ma_cnt = 1 To q_data(a_index).ct_count
                                    mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                                Next ma_cnt
                                If mna_cnt = 0 Then
                                    na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                                End If
                            End If
                        End If
                    Else
                        If weight_flg = "�Ȃ�" Then
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + 1
                            End If
                        Else
                            For ma_cnt = 1 To q_data(a_index).ct_count
                                mna_cnt = mna_cnt + Val(ws_outdata.Cells(i_cnt, q_data(a_index).data_column + ma_cnt - 1))
                            Next ma_cnt
                            If mna_cnt = 0 Then
                                na_cnt = na_cnt + Val(ws_outdata.Cells(i_cnt, q_data(w_index).data_column))
                            End If
                        End If
                    End If
                End If
            End If
        Next i_cnt
    End If
    
    ' �m�L���񓚐��n�̎Z�o
    vr_cnt = face_cnt - na_cnt
    
    ' �����m�\�����ځn�̏���
    ws_summary1.Cells(sum_row, sum_col - 1).Value = "N/A"
    ws_summary2.Cells(div_row, div_col - 1).Value = "N/A"
    ws_summary3.Cells(div_row, div_col - 1).Value = "N/A"
    ws_summary1.Cells(sum_row, sum_col + 1).Value = "����"
    ws_summary2.Cells(div_row, div_col + 1).Value = "����"
    ws_summary3.Cells(div_row, div_col + 1).Value = "����"
    sum_col = sum_col + 2
    div_col = div_col + 2
    
    If c_data(c_cnt).ken_flg = "Y" Then
        ' �������Ɂm�L���񓚐��n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        ws_summary3.Cells(div_row, div_col).Value = vr_cnt
    Else
        ' �������Ɂm�\�����ڑS���n���o��
        ws_summary1.Cells(sum_row, sum_col).Value = face_cnt
        ws_summary2.Cells(div_row, div_col).Value = face_cnt
        ws_summary3.Cells(div_row, div_col).Value = face_cnt
    End If
    
    ' �E�G�C�g�W�v���m�S�́n�̃Z�������ݒ�
    If weight_flg = "����" Then
        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
    End If
    sum_col = sum_col + 1
    div_col = div_col + 1
    
    ' �����m�J�e�S���[�n�̏���
    total_cnt = 0
    If ama_cnt = 0 Then
        For a_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                ' �\���l�`�̂m�`�������\���r�`
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"

                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column), ws_outdata.Cells(outdata_row, q_data(a_index).data_column)), a_cnt, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next a_cnt
    Else
        For ma_cnt = 1 To q_data(a_index).ct_count
            If fma_cnt = 0 Then
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "", _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, q_data(f_index).data_column), ws_outdata.Cells(outdata_row, q_data(f_index).data_column)), "")
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            Else
                ' �\���l�`�̂m�`�������\���l�`
                If select_flg = 1 Then
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1, _
                         Range(ws_outdata.Cells(7, 16384), ws_outdata.Cells(outdata_row, 16384)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                Else
                    If weight_flg = "�Ȃ�" Then
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         CountIfs(Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                    Else
                        ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            
                        ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction. _
                         SumIfs(Range(ws_outdata.Cells(7, q_data(w_index).data_column), ws_outdata.Cells(outdata_row, q_data(w_index).data_column)), _
                         Range(ws_outdata.Cells(7, q_data(a_index).data_column + ma_cnt - 1), ws_outdata.Cells(outdata_row, q_data(a_index).data_column + ma_cnt - 1)), 1, _
                         Range(ws_outdata.Cells(7, 16383), ws_outdata.Cells(outdata_row, 16383)), 1)
                        ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
                    End If
                End If
            End If
            
            '�\����̎Z�o
            item_cnt = ws_summary1.Cells(sum_row, sum_col)
            If c_data(c_cnt).bosu_flg = "Y" Then
                If vr_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
               Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / vr_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            Else
                If face_cnt = 0 Then
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                    ws_summary3.Cells(div_row, div_col).Value = "-"
                Else
                    ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                    ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(item_cnt / face_cnt * 100, 1)
                    ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
                End If
            End If
            total_cnt = total_cnt + item_cnt
            sum_col = sum_col + 1
            div_col = div_col + 1
        Next ma_cnt
    End If
    
    ' �����m���񓚁n�̏���
    If c_data(c_cnt).ana_flg <> "N" Then
        ws_summary1.Cells(sum_row, sum_col).Value = na_cnt
        ws_summary2.Cells(div_row, div_col).Value = na_cnt
        
        ' �E�G�C�g�W�v���m���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If face_cnt = 0 Then
            ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        Else
            ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
            ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(na_cnt / face_cnt * 100, 1)
            ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m�L���񓚁n�̏���
    If c_data(c_cnt).yuko_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = vr_cnt
        ws_summary2.Cells(div_row, div_col).Value = vr_cnt
        
        ' �E�G�C�g�W�v���m�L���񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(vr_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����m���׉񓚁n�̏���
    If c_data(c_cnt).nobe_flg = "Y" Then
        ws_summary1.Cells(sum_row, sum_col).Value = total_cnt
        ws_summary2.Cells(div_row, div_col).Value = total_cnt
        
        ' �E�G�C�g�W�v���m���׉񓚁n�̃Z�������ݒ�
        If weight_flg = "����" Then
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
            ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        End If
        
        If c_data(c_cnt).bosu_flg = "Y" Then
            If vr_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / vr_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        Else
            If face_cnt = 0 Then
                ws_summary1.Cells(sum_row + 1, sum_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row + 1, sum_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary1.Cells(sum_row + 1, sum_col).NumberFormatLocal = "0.0"
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.Round(total_cnt / face_cnt * 100, 1)
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0.0"
            End If
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �W�v�ݒ�t�@�C���̏�������m�����ݖ�p�n
    wb_outdata.Activate
    ws_outdata.Select
    filter_flg = 0
    If fma_cnt = 0 Then
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, "=", visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter q_data(f_index).data_column, "=", visibledropdown:=False
        End If
    Else
        ' �\���l�`�̂m�`�����������ݖ�
        If select_flg = 1 Then
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16384, 1, visibledropdown:=False
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16383, 1, visibledropdown:=False
        Else
            Range(ws_outdata.Cells(6, 1), ws_outdata.Cells(outdata_row, 16384)).AutoFilter 16383, 1, visibledropdown:=False
        End If
    End If
    
    If WorksheetFunction.Subtotal(3, Range(ws_outdata.Cells(7, 1), ws_outdata.Cells(outdata_row, 1))) = 0 Then
        filter_flg = 1
    Else
        filter_flg = 99
    End If
    
    If weight_flg = "�Ȃ�" Then
        Call Real_Answer(filter_flg, vr_cnt)    ' �e�����ݖ�̏�����
    Else
        Call Real_Answer_WGT(filter_flg, vr_cnt)    ' �e�����ݖ�i�E�G�C�g�W�v�j�̏�����
    End If
    
    '�I�[�g�t�B���^�̉���
    wb_outdata.Activate
    ws_outdata.Select
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.ShowAllData
    End If
    
    sum_col = temp_col
    div_col = temp_col
End Sub

Private Sub Real_Answer(ByVal f_flag As Integer, ByVal v_cnt As Long)
    On Error Resume Next
    Dim decimal_places As Integer
' �e�����ݖ�̏���
    ' �����ݖ�m���v�n�̏���
    If c_data(c_cnt).sum_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                   
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m���ρn�̏���
    decimal_places = Val(c_data(c_cnt).ave_flg)
    If c_data(c_cnt).ave_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), 0), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�W���΍��n�̏���
    decimal_places = Val(c_data(c_cnt).sd_flg)
    If c_data(c_cnt).sd_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ŏ��l�n�̏���
    If c_data(c_cnt).min_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m��P�l���ʁn�̏���
    If c_data(c_cnt).q1_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 1), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�����l�n�̏���
    If c_data(c_cnt).med_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m��R�l���ʁn�̏���
    If c_data(c_cnt).q3_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column)), 3), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ő�l�n�̏���
    If c_data(c_cnt).max_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ŕp�l�n�̏���
    If c_data(c_cnt).mod_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, q_data(ra_index).data_column), ws_outdata.Cells(outdata_row, q_data(ra_index).data_column))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
End Sub

Private Sub Real_Answer_WGT(ByVal f_flag As Integer, ByVal v_cnt As Long)
    On Error Resume Next
    Dim decimal_places As Integer
' �e�����ݖ�i�E�G�C�g�W�v�j�̏���
    ' �����ݖ�m���v�n�̏���
    If c_data(c_cnt).sum_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0"
                    
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0"
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(9, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0"
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m���ρn�̏���
    decimal_places = Val(c_data(c_cnt).ave_flg)
    If c_data(c_cnt).ave_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), 0), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).ave_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(1, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�W���΍��n�̏���
    decimal_places = Val(c_data(c_cnt).sd_flg)
    If c_data(c_cnt).sd_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg = "0" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    ElseIf c_data(c_cnt).sd_flg <> "" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
             Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
            ws_summary1.Cells(sum_row, sum_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary2.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError(Application.WorksheetFunction.Round( _
                 Application.WorksheetFunction.Aggregate(8, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), decimal_places), "-")
                ws_summary3.Cells(div_row, div_col).NumberFormatLocal = "0." & String(decimal_places, "0")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ŏ��l�n�̏���
    If c_data(c_cnt).min_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(5, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m��P�l���ʁn�̏���
    If c_data(c_cnt).q1_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
            Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 1), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�����l�n�̏���
    If c_data(c_cnt).med_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(12, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m��R�l���ʁn�̏���
    If c_data(c_cnt).q3_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(17, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382)), 3), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ő�l�n�̏���
    If c_data(c_cnt).max_flg = "Y" Then
        If f_flag = 0 Then
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            If v_cnt = 0 Then    ' �L���񓚃[������
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
                
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(4, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
    
    ' �����ݖ�m�ŕp�l�n�̏���
    If c_data(c_cnt).mod_flg = "Y" Then
        If f_flag = 0 Then
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        ElseIf f_flag = 99 Then    ' �\�����񓚐�p����
            ws_summary1.Cells(sum_row, sum_col).Value = Application.WorksheetFunction.IfError( _
             Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
            If ws_summary1.Cells(sum_row, sum_col).Value = "" Then
                ws_summary1.Cells(sum_row, sum_col).Value = "-"
                ws_summary2.Cells(div_row, div_col).Value = "-"
                ws_summary3.Cells(div_row, div_col).Value = "-"
            Else
                ws_summary2.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
        
                ws_summary3.Cells(div_row, div_col).Value = Application.WorksheetFunction.IfError( _
                 Application.WorksheetFunction.Aggregate(13, 7, Range(ws_outdata.Cells(7, 16382), ws_outdata.Cells(outdata_row, 16382))), "-")
            End If
        Else
            ws_summary1.Cells(sum_row, sum_col).Value = "-"
            ws_summary2.Cells(div_row, div_col).Value = "-"
            ws_summary3.Cells(div_row, div_col).Value = "-"
        End If
        sum_col = sum_col + 1
        div_col = div_col + 1
    End If
End Sub

Private Sub weight_ra()
    Dim i_cnt As Long
' �W�v�Ώۃf�[�^���Ƃ̃E�G�C�g�W�v�p�����l�̎Z�o
' �y�T�v�z�W�v�Ώۃf�[�^�ƂȂ�t�@�C���́m16382�n��ڂɁA�����l�ƃE�G�C�g�i�␳�l�j�̐ς��o�́B
'
    If ra_index <> 0 Then    ' �\�����ڂɎ����ݖ�̎w�肪����Ώ�������B
        wb_outdata.Activate
        ws_outdata.Select
        Columns("XFB:XFB").Select
        Selection.ClearContents
        Cells(6, 16382) = "weight_ra"
    
        ' �T���v���x�[�X�ł̎����l�~�E�G�C�g�i�␳�l�j�̎Z�o
        For i_cnt = 7 To outdata_row
            If ws_outdata.Cells(i_cnt, q_data(ra_index).data_column) <> "" Then
                ws_outdata.Cells(i_cnt, 16382) = _
                 ws_outdata.Cells(i_cnt, q_data(ra_index).data_column) * ws_outdata.Cells(i_cnt, q_data(w_index).data_column)
            End If
        Next i_cnt
    End If
End Sub

Private Sub Mcode_Setting()
' MCODE���� - 2020.1.10 �ǉ��A2020.3.26 �ҏW
    Dim i_cnt As Long, m_cnt As Long
    Dim max_row As Long
    Dim s_pos As Long
    Dim m_code As String
    Dim hyo_num As String
    Dim bgn_row As Long, fin_row As Long
    Dim head_cm As String, face_cm As String
    
'�y�m���\�z
    wb_summary.Activate
    ws_summary1.Select
    
    ' �T�}���[�t�@�C���̍ŏI�s�擾�iG��Ŏ擾���Ă܂��j
    max_row = ws_summary1.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODE�̌����iB�񂾂��ł͂Ȃ��A�\�ԍ��Ƃ��킹�Č����j
        If ws_summary1.Cells(i_cnt, 1) <> "" Then
            If ws_summary1.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "�^")
                    head_cm = Left(ws_summary1.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary1.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' �P���W�v�̔��� - 2020.3.26
                    If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                      ws_summary1.Cells(i_cnt, 4) = head_cm
                      ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary1.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary1.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary1.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "�^")
                        face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary1.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' �P���W�v�̔��� - 2020.3.26
                        If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                          ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary1.Cells(ActiveCell.Row, 4).Clear
                          ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary1.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary1.Cells(i_cnt, 2)
                        m_cnt = 1
                    
                        s_pos = InStr(ws_summary1.Cells(i_cnt, 4), "�^")
                        head_cm = Left(ws_summary1.Cells(i_cnt, 4), s_pos - 1)
                        face_cm = Mid(ws_summary1.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary1.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' �P���W�v�̔��� - 2020.3.26
                        If ws_summary1.Cells(ActiveCell.Row + 2, 3) = "" Then
                          ws_summary1.Cells(i_cnt, 4) = head_cm
                          ws_summary1.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary1.Cells(ActiveCell.Row, 4) = head_cm
                          ws_summary1.Cells(ActiveCell.Row, 5) = face_cm
                          m_code = ws_summary1.Cells(i_cnt, 2)
                          m_cnt = m_cnt + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i_cnt

'�y�m�\�z
    wb_summary.Activate
    ws_summary2.Select
    
    ' �T�}���[�t�@�C���̍ŏI�s�擾�iG��Ŏ擾���Ă܂��j
    max_row = ws_summary2.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODE�̌����iB�񂾂��ł͂Ȃ��A�\�ԍ��Ƃ��킹�Č����j
        If ws_summary2.Cells(i_cnt, 1) <> "" Then
            If ws_summary2.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary2.Cells(i_cnt, 4), "�^")
                    head_cm = Left(ws_summary2.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary2.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary2.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' �P���W�v�̔��� - 2020.3.26
                    If ws_summary2.Cells(ActiveCell.Row + 1, 3) = "" Then
                      ws_summary2.Cells(i_cnt, 4) = head_cm
                      ws_summary2.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary2.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary2.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary2.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary2.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary2.Cells(i_cnt, 4), "�^")
                        face_cm = Mid(ws_summary2.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary2.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' �P���W�v�̔��� - 2020.3.26
                        If ws_summary2.Cells(ActiveCell.Row + 1, 3) = "" Then
                          ws_summary2.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary2.Cells(ActiveCell.Row, 4).Clear
                          ws_summary2.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary2.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary2.Cells(i_cnt, 2)
                        m_cnt = 1
                    End If
                End If
            End If
        End If
    Next i_cnt

'�y���\�z
    wb_summary.Activate
    ws_summary3.Select
    
    ' �T�}���[�t�@�C���̍ŏI�s�擾�iG��Ŏ擾���Ă܂��j
    max_row = ws_summary3.Cells(Rows.Count, 7).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 1 To max_row
        ' MCODE�̌����iB�񂾂��ł͂Ȃ��A�\�ԍ��Ƃ��킹�Č����j
        If ws_summary3.Cells(i_cnt, 1) <> "" Then
            If ws_summary3.Cells(i_cnt, 2) <> "" Then
                If m_cnt = 1 Then
                    s_pos = InStr(ws_summary3.Cells(i_cnt, 4), "�^")
                    head_cm = Left(ws_summary3.Cells(i_cnt, 4), s_pos - 1)
                    face_cm = Mid(ws_summary3.Cells(i_cnt, 4), s_pos + 1)
                    ws_summary3.Cells(i_cnt, 3).Select
                    Selection.End(xlDown).Select
                    ' �P���W�v�̔��� - 2020.3.26
                    If ws_summary3.Cells(ActiveCell.Row + 1, 3) = "" Then
                      ws_summary3.Cells(i_cnt, 4) = head_cm
                      ws_summary3.Cells(ActiveCell.Row, 3) = m_cnt
                      ws_summary3.Cells(ActiveCell.Row, 4) = head_cm
                      ws_summary3.Cells(ActiveCell.Row, 5) = face_cm
                      m_code = ws_summary3.Cells(i_cnt, 2)
                      m_cnt = m_cnt + 1
                    End If
                Else
                    If ws_summary3.Cells(i_cnt, 2) = m_code Then
                        bgn_row = i_cnt - 1
                        s_pos = InStr(ws_summary3.Cells(i_cnt, 4), "�^")
                        face_cm = Mid(ws_summary3.Cells(i_cnt, 4), s_pos + 1)
                        ws_summary3.Cells(i_cnt, 3).Select
                        Selection.End(xlDown).Select
                        ' �P���W�v�̔��� - 2020.3.26
                        If ws_summary3.Cells(ActiveCell.Row + 1, 3) = "" Then
                          ws_summary3.Cells(ActiveCell.Row, 3) = m_cnt
                          ws_summary3.Cells(ActiveCell.Row, 4).Clear
                          ws_summary3.Cells(ActiveCell.Row, 5) = face_cm
                          fin_row = ActiveCell.Row - 1
                          ws_summary3.Range(bgn_row & ":" & fin_row).Delete
                          m_cnt = m_cnt + 1
                        End If
                    Else
                        m_code = ws_summary3.Cells(i_cnt, 2)
                        m_cnt = 1
                    End If
                End If
            End If
        End If
    Next i_cnt

'�y�ڎ��z
    wb_summary.Activate
    ws_summary0.Select
    
    ' �T�}���[�t�@�C���̍ŏI�s�擾�iA��Ŏ擾���Ă܂��j
    max_row = ws_summary0.Cells(Rows.Count, 1).End(xlUp).Row
    
    m_cnt = 1
    m_code = ""
    For i_cnt = 2 To max_row    ' �w�b�_�[������̂ŊJ�n��[2]����B
        ' MCODE�̌����iB�񂾂��ł͂Ȃ��A�\�ԍ��Ƃ��킹�Č����j
        If ws_summary0.Cells(i_cnt, 2) <> "" Then
            If ws_summary0.Cells(i_cnt, 3) <> "" Then
                If ws_summary0.Cells(i_cnt, 4) = "" Then
                    If m_cnt = 1 Then
                        hyo_num = ws_summary0.Cells(i_cnt, 2)
                        m_code = ws_summary0.Cells(i_cnt, 3)
                        m_cnt = m_cnt + 1
                    Else
                        If ws_summary0.Cells(i_cnt, 3) = m_code Then
                            ws_summary0.Cells(i_cnt, 2) = "'" & hyo_num
                            m_cnt = m_cnt + 1
                        Else
                            hyo_num = ws_summary0.Cells(i_cnt, 2)
                            m_code = ws_summary0.Cells(i_cnt, 3)
                            m_cnt = 1
                        End If
                    End If
                End If
            End If
        End If
    Next i_cnt
End Sub
