Attribute VB_Name = "Module00"
Option Explicit

Public incol_array() As String
Public wb            As Workbook
Public wb_indata     As Workbook
Public ws_indata     As Worksheet
Public wb_outdata    As Workbook
Public ws_outdata    As Worksheet
Public ws_mainmenu   As Worksheet
Public ws_setup      As Worksheet

Public gcode_row, gcode_col     As Integer
Public gdrive_row, gdrive_col   As Integer
Public initial_row, initial_col As Integer
Public setup_row, setup_col     As Integer
Public incol_row, incol_col     As Integer

Public file_path    As String
Public ope_code     As String
Public outdata_fn   As String
Public status_msg   As String
Public progress_msg As String

' 20170419 ���R�� �����pQCODE�z��
Public str_code() As String

' 20170425 ���R�� �ݒ��ʊi�[�p���[�U�^
Public Type question_data
q_code As String                ' QCODE�i�[�p������ϐ�
r_code As String                ' ����CODE�i�[�p������ϐ�
m_code As String                ' MCODE�i�[�p������ϐ�
q_format As String              ' �ݖ�`���i�[�p������ϐ�
r_byte As Integer               ' ���������i�[�p�ϐ�
r_unit As String                ' �����P�ʕ�����ϐ��i2017.05.10 �ǉ��j
sel_code1 As String             ' �Z���N�g�����@QCODE�i�[�p������ϐ�
sel_value1 As Integer           ' �Z���N�g�����@�l�i�[�p������ϐ�
sel_code2 As String             ' �Z���N�g�����AQCODE�i�[�p������ϐ�
sel_value2 As Integer           ' �Z���N�g�����A�l�i�[�p������ϐ�
sel_code3 As String             ' �Z���N�g�����BQCODE�i�[�p������ϐ�
sel_value3 As Integer           ' �Z���N�g�����B�l�i�[�p������ϐ�
ct_count As Integer             ' �ݖ�J�e�S���[���i�[�p�ϐ�
ct_loop As Integer              ' ���[�v�J�E���g���i�[�p�ϐ�
'        ct_0flg As Boolean              ' �O�J�e�S���[�t���O�i�[�p�ϐ�
q_title As String               ' �\��i�[�p������ϐ�
q_ct(300) As String             ' �ݖ⎈�i�[�p������z��
data_column As Integer          ' ���̓f�[�^��ԍ��i�[�p�ϐ�
End Type

Public q_data() As question_data    ' QCODE���̃f�[�^��S�Ď擾

public Sub Auto_Open()
  Dim base_pt As String
  '--------------------------------------------------------------------------------------------------'
  '�@�t�@�C���I�[�v�������@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.10�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03  '
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  '    Call Starting_Mcs2017    ' ThisWorkbook �ɂČĂяo���ɕύX - 2020.6.3
  Application.StatusBar = "�t�@�C���I�[�v��������..."
  Application.ScreenUpdating = False
  wb.Activate
  ws_mainmenu.Select
  If ws_mainmenu.Cells(gdrive_row, gdrive_col) <> "" And _
    ws_mainmenu.Cells(gcode_row, gcode_col) <> "" Then
    base_pt = ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If Dir(base_pt & "\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
      If (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 8) <> "// �ǂݍ��񂾓���") And _
        (Mid(ws_mainmenu.Cells(initial_row, initial_col), 1, 7) <> "// �ۑ���������") Then
        ws_mainmenu.Cells(initial_row, initial_col) = "// �����ݒ�ς�"
      End If
    Else
      ws_mainmenu.Cells(initial_row, initial_col) = ""
    End If
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
  End If
  Application.ScreenUpdating = True
  Application.StatusBar = False
End Sub
'-----------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------'
'�@�V�[�g�\���̏����ݒ�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.10�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2017.04.14�@'
'--------------------------------------------------------------------------------------------------'
public Sub Starting_Mcs2017()
  Set wb = ThisWorkbook
  Set ws_mainmenu = wb.Worksheets("���C�����j���[")
  Set ws_setup = wb.Worksheets("�ݒ���")
  ' ���C�����j���[�F�Ɩ��R�[�h�̍s��
  gcode_row = 3
  gcode_col = 8
  ' ���C�����j���[�F��ƃh���C�u�̍s��
  gdrive_row = 3
  gdrive_col = 23
  ' ���C�����j���[�F�����ݒ�ς݃��b�Z�[�W�o�͐�s��
  initial_row = 6
  initial_col = 32
  ' �ݒ��ʁF�擪�p�����[�^�̍s��
  setup_row = 3
  setup_col = 1
  ' �ݒ��ʁF���̓f�[�^��ԍ��̍s��
  incol_row = 3
  incol_col = 3
End Sub
'-----------------------------------------------------------------------------

public Sub Finishing_Mcs2017()
  '--------------------------------------------------------------------------------------------------'
  '�@�I�������E�e�I�u�W�F�N�g�̎Q�Ƃ������@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.10�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2018.09.19�@'
  '--------------------------------------------------------------------------------------------------'
  Application.ScreenUpdating = True
  Application.StatusBar = False

  wb.Activate
  ws_setup.Select
  ws_setup.Cells(3, 1).Select
  ws_mainmenu.Select
  ws_mainmenu.Cells(3, 8).Select

  ' �ݒ��ʃ`�F�b�N�̃G���[���O�t�@�C����0�o�C�g�Ȃ�t�@�C���폜
  If Dir(file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx") <> "" Then
    If FileLen(file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx") = 0 Then
      Kill file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx"
    End If
  End If

  ' ���W�b�N�`�F�b�N�ɂ��f�[�^�t�@�C���̃G���[���O�t�@�C����0�o�C�g�Ȃ�t�@�C���폜
  If Dir(file_path & "\4_LOG\" & ope_code & "err.xlsx") <> "" Then
    If FileLen(file_path & "\4_LOG\" & ope_code & "err.xlsx") = 0 Then
      Kill file_path & "\4_LOG\" & ope_code & "err.xlsx"
    End If
  End If

  Application.DisplayAlerts = False
  wb.Save
  Application.DisplayAlerts = True

  Set wb = Nothing
  Set ws_mainmenu = Nothing
  Set ws_setup = Nothing
End Sub
'-----------------------------------------------------------------------------

public Sub Setup_Check()
  Dim qcode_cnt As Long
  Dim setup_status As String

  Dim wb_preset As Workbook
  Dim ws_preset As Worksheet
  Dim preset_row As Variant
  Dim preset_col As Variant
  Dim preset_i As Long
  Dim preset_j As Long
  Dim preset_gcode As String

  Dim max_row As Long
  Dim max_col As Long
  Dim err_msg As String
  Dim err_cnt As Long
  Dim sel_i As Long
  Dim selcode_row As Long
  Dim sys_cnt As Long
  Dim sys_num As Long
  Dim sys_col As Long  '5,6,7
  Dim sys_i As Long
  Dim mcode_col As Long
  Dim qCStr_col As Long
  Dim qformat_typ As String
  Dim selcode_col_1 As Long
  Dim selval_col_1 As Long
  Dim selcode_col_2 As Long
  Dim selval_col_2 As Long
  Dim selcode_col_3 As Long
  Dim selval_col_3 As Long
  Dim find_row_1 As Long
  Dim find_row_2 As Long
  Dim find_row_3 As Long
  Dim ct_cnt_col As Long
  Dim zero_f_col As Long
  Dim qttl_col As Long
  Dim ct_st_col As Long
  Dim rcode_col As Long
  Dim rcode_row As Long
  Dim rcode_cell As Range
  '--------------------------------------------------------------------------------------------------'
  '�@�ݒ��ʃ`�F�b�N�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.12�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.05.19�@'
  '--------------------------------------------------------------------------------------------------'
  On Error Resume Next
  Call Starting_Mcs2017
  Application.StatusBar = "�ݒ��� �`�F�b�N��..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If ws_mainmenu.Cells(gcode_row, gcode_col).Value = "" Then
    MsgBox "���C�����j���[�̋Ɩ��R�[�h�������͂ł��B", vbExclamation, "MCS 2020 - Setup_Check"
    Cells(gcode_row, gcode_col).Select
    Application.StatusBar = False
    End
  End If

  Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & _
  "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col).Value & "_mcs.ini" For Input As #1
  Line Input #1, file_path
  Line Input #1, setup_status
  Close #1

  Open file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx" For Append As #1
  Close #1
  If Err.Number > 0 Then
    Workbooks(ope_code & "_�ݒ���NG.xlsx").Close
  End If

  If Dir(file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx") <> "" Then
    Kill file_path & "\4_LOG\" & ope_code & "_�ݒ���NG.xlsx"
  End If

  ws_setup.Select
  max_row = Cells(Rows.Count, setup_col).End(xlUp).Row

  ' ��������G���[�`�F�b�N�̃R�[�f�B���O(�L��֥`)
  err_msg = "�y�ݒ��ʂ̃G���[���X�g�z" & vbCrLf & "QCODE,�s��,�G���[����,�G���[���e" & vbCrLf

  '���T���v���i���o�[��QCODE�`�F�b�N
  If Cells(3, 1).Value <> "SNO" Then
    err_msg = err_msg & Cells(3, 1).Value & ",3�s��" & ",QCODE(RC1)" & ",�T���v������QCODE�́mSNO�n�Ƃ��Ă��������B" & vbCrLf
  End If

  For qcode_cnt = setup_row To max_row
    If Cells(qcode_cnt, setup_col).Value = "*���H��" Then
      Rows(qcode_cnt).Select
      With Selection.Interior
        .Color = 65535
      End With
    ElseIf Left(Cells(qcode_cnt, setup_col).Value, 1) = "*" Then
      '��������
    Else
      '�ݒ��ʍs����o�^
      max_col = Cells(qcode_cnt, Columns.Count).End(xlToLeft).Column
      sys_cnt = 3: sys_num = 5
      mcode_col = 8: qCStr_col = 9
      selcode_col_1 = 10: selval_col_1 = 11
      selcode_col_2 = 12: selval_col_2 = 13
      selcode_col_3 = 14: selval_col_3 = 15
      ct_cnt_col = 16: zero_f_col = 17
      qttl_col = 18: ct_st_col = 19
      find_row_1 = 0
      find_row_2 = 0
      find_row_3 = 0
      rcode_col = 2: rcode_row = 0

    '��QCODE �̏d���`�F�b�N
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
    Cells(qcode_cnt, setup_col).Value)
    If err_cnt >= 2 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",QCODE(RC1)" & ",QCODE�͏d���s�ł��B" & vbCrLf
    End If

    '��QCODE���w�肩��R�����g�s�ւ̓��̓`�F�b�N
    If Cells(qcode_cnt, setup_col).Value = "" And _
      WorksheetFunction.CountA(Range(Cells(qcode_cnt, setup_col), Cells(qcode_cnt, max_col))) <> 0 Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",QCODE(RC1)" & ",QCODE���u�����N�̏ꍇ�A���̐ݒ荀�ڂւ̓��͕͂s�ƂȂ�܂��B" & vbCrLf
    End If

    '��MCODE�P�Ɠ��̓`�F�b�N
    err_cnt = WorksheetFunction.CountIf(Range(Cells(setup_row, mcode_col), Cells(max_row, mcode_col)), _
    Cells(qcode_cnt, mcode_col).Value)
    If err_cnt = 1 And Cells(qcode_cnt, mcode_col).Value <> "" Then
      err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",MCODE(RC8)" & ",MCODE��2�ȏ�̐ݖ�ɑ΂��Ďw�肵�Ă��������B" & vbCrLf
    End If

    '���`���`�F�b�N
    qformat_typ = Cells(qcode_cnt, qCStr_col).Value
    Select Case Left(qformat_typ, 1)
      Case "C", "S", "M", "H", "F", "O"
        If Len(qformat_typ) > 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�`����1���ŁmC�n�mS�n�mM�n�mH�n�mF�n�mO�n���w�肵�Ă��������B" & vbCrLf
        End If
      Case "R"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�`���mR�n�͌������w�肵�Ă��������B" & vbCrLf
        ElseIf Len(qformat_typ) > 3 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",������2���܂Ŏw��\�ł��B" & vbCrLf
        ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�����͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
        End If
      Case "L"
        If Len(qformat_typ) = 1 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�`�� �mL�n�͉񓚐������̎w������Ă��������B" & vbCrLf
        ElseIf (Mid(qformat_typ, 2, 1) = "M") Or (Mid(qformat_typ, 2, 1) = "A") Or (Mid(qformat_typ, 2, 1) = "C") Then
          If Len(qformat_typ) = 2 Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�`�� �mL�n�͉񓚐������̎w������Ă��������B" & vbCrLf
          ElseIf Len(qformat_typ) > 5 And IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�񓚐�������3���܂Ŏw��\�ł��B" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 3, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�񓚐������͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
          End If
        ElseIf Val(Mid(qformat_typ, 2, 1)) = 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",���~�b�g�}���`�̌`���́mLMx�n�mLAx�n�mLCx�n�̂����ꂩ�Ŏw������Ă��������ix�͉񓚐��j�B" & vbCrLf
        Else
          If Len(qformat_typ) > 4 And IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = True Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�񓚐�������3���܂Ŏw��\�ł��B" & vbCrLf
          ElseIf IsNumeric(Val(Mid(qformat_typ, 2, Len(qformat_typ) - 1))) = False Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�񓚐������͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
          End If
        End If
      Case Else
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�`��(RC9)" & ",�`���mC�n�mS�n�mM�n�mH�n�mF�n�mO�n�mR�n�mL�n�ȊO�̎w��͕s�ƂȂ�܂��B" & vbCrLf
    End Select

    '���Z���N�g�`�F�b�N
    err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3)))
    If err_cnt <> 0 Then
      '���Z���N�g�����L�@�`�F�b�N
      If WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_1))) < 2 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_2), Cells(qcode_cnt, selval_col_2))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@�A�B(RC10:RC15)" & ",�����i�Z���N�g�j�̎w���QCODE�ƒl�̃Z�b�g�������@���獶�l�߂Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_2))) < 4 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@�A�B(RC10:RC15)" & ",�����i�Z���N�g�j�̎w���QCODE�ƒl�̃Z�b�g�������@���獶�l�߂Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_1), Cells(qcode_cnt, selval_col_3))) < 6 And _
        WorksheetFunction.CountA(Range(Cells(qcode_cnt, selcode_col_3), Cells(qcode_cnt, selval_col_3))) >= 1 Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@�A�B(RC10:RC15)" & ",�����i�Z���N�g�j�̎w���QCODE�ƒl�̃Z�b�g�������@���獶�l�߂Ŏw�肵�Ă��������B" & vbCrLf
      End If

      '���Z���N�g����QCODE�Y���`�F�b�N
      '�����P
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_1).Value) = 0 And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����i�Z���N�g�j�ƂȂ�QCODE�͗\��A��Ɏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        find_row_1 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_1).Value, lookat:=xlWhole).Row
      End If
      '����2
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_2).Value) = 0 And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����i�Z���N�g�j�ƂȂ�QCODE�͗\��A��Ɏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        find_row_2 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_2).Value, lookat:=xlWhole).Row
      End If
      '����3
      If WorksheetFunction.CountIf(Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)), _
        Cells(qcode_cnt, selcode_col_3).Value) = 0 And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����i�Z���N�g�j�ƂȂ�QCODE�͗\��A��Ɏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        find_row_3 = Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
        Find(Cells(qcode_cnt, selcode_col_3).Value, lookat:=xlWhole).Row
      End If

      '���Z���N�g�����l�`�F�b�N
      '�����P
      If Cells(qcode_cnt, selval_col_1).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����i�Z���N�g�j�̒l�͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value = "" And Cells(qcode_cnt, selcode_col_1).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_1).Value <> "" And Cells(qcode_cnt, selcode_col_1).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_1).Value) = True And find_row_1 <> 0 Then
        If Cells(find_row_1, ct_cnt_col).Value < Val(Cells(qcode_cnt, selval_col_1).Value) Then
          Debug.Print find_row_1
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����i�Z���N�g�j�̒l�͊Y���ݖ�̑I�������ȉ��̐��l�Ŏw�肵�Ă��������B" & vbCrLf
        End If
      End If
      '�����Q
      If Cells(qcode_cnt, selval_col_2).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����i�Z���N�g�j�̒l�͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value = "" And Cells(qcode_cnt, selcode_col_2).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_2).Value <> "" And Cells(qcode_cnt, selcode_col_2).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_2).Value) = True And find_row_2 <> 0 Then
        If Cells(find_row_2, ct_cnt_col).Value - Cells(find_row_2, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_2).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����i�Z���N�g�j�̒l�͊Y���ݖ�̑I�������ȉ��̐��l�Ŏw�肵�Ă��������B" & vbCrLf
        End If
      End If
      '�����R
      If Cells(qcode_cnt, selval_col_3).Value <> "" And IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = False Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����i�Z���N�g�j�̒l�͔��p�����Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value = "" And Cells(qcode_cnt, selcode_col_3).Value <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf Cells(qcode_cnt, selval_col_3).Value <> "" And Cells(qcode_cnt, selcode_col_3).Value = "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����i�Z���N�g�j��QCODE�ƒl�̃Z�b�g�Ŏw�肵�Ă��������B" & vbCrLf
      ElseIf IsNumeric(Cells(qcode_cnt, selval_col_3).Value) = True And find_row_3 <> 0 Then
        If Cells(find_row_3, ct_cnt_col).Value - Cells(find_row_3, zero_f_col).Value _
          < Val(Cells(qcode_cnt, selval_col_3).Value) Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����i�Z���N�g�j�̒l�͊Y���ݖ�̑I�������ȉ��̐��l�Ŏw�肵�Ă��������B" & vbCrLf
        End If
      End If
    End If

    '���I�������`�F�b�N
    Select Case Left(Cells(qcode_cnt, 9).Value, 1)
      Case "S", "M", "L"
        If max_col >= ct_st_col Then
          err_cnt = WorksheetFunction.CountA(Range(Cells(qcode_cnt, ct_st_col), Cells(qcode_cnt, max_col)))
          If err_cnt <> Cells(qcode_cnt, ct_cnt_col).Value Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",CT���E�J�e�S���[���e(RC16��RC19�`)" & ",P���CT���ƁAS��ȍ~�̃J�e�S���[���ڂ̌�����v���܂���B" & vbCrLf
          ElseIf err_cnt <> max_col - qttl_col Then
            err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�J�e�S���[����(RC19�`)" & ",S��ȍ~�̃J�e�S���[���ڂ͍��l�߂œ��͂��Ă��������B" & vbCrLf
          End If
        End If
      Case "F", "O", "R", "H", "C"
        If max_col >= ct_st_col Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�J�e�S���[����(RC19�`)" & ",�`���mR�n�mH�n�mF�n�mO�n�mC�n�̃J�e�S���[���ڂ̐ݒ�͂ł��܂���B" & vbCrLf
        ElseIf Cells(qcode_cnt, ct_cnt_col).Value <> 0 Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",CT��(RC16)" & ",�`���mR�n�mH�n�mF�n�mO�n�mC�n�̐ݖ��CT���́m�u�����N�n�ɂ��Ă��������B" & vbCrLf
        End If
      Case Else
    End Select

    ' 2020.1.9 - �[���t���O�͂Ȃ��Ȃ�܂����̂ŁA�R�����g�A�E�g���܂����B
    '            '���[���t���O�`�F�b�N
    '            If Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Len(Cells(qcode_cnt, ct_cnt_col).Value) = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",0CT(RC17)" & ",0CT�̃t���O��CT�����m1�ȏ�n�̐ݖ�ɂ̂ݐݒ�\�ł��B" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, ct_cnt_col).Value = 0 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",0CT(RC17)" & ",0CT�̃t���O��CT�����m1�ȏ�n�̐ݖ�ɂ̂ݐݒ�\�ł��B" & vbCrLf
    '            ElseIf Len(Cells(qcode_cnt, zero_f_col).Value) <> 0 And Cells(qcode_cnt, zero_f_col).Value <> 1 Then
    '                err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",0CT(RC17)" & ",0CT�̎w��Ƃ��Ďg����̂́m1�n�݂̂ƂȂ�܂��B" & vbCrLf
    '            End If

    '���ݖ�^�C�g���L���`�F�b�N
    If Cells(qcode_cnt, qttl_col) = "" Then
      If Cells(qcode_cnt, ct_st_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�\��(RC18)" & ",�ݖ�̕\��͕K���w�肵�Ă��������B" & vbCrLf
      End If
    End If

    '���\���G���A�ւ̓��̓`�F�b�N
    For sys_i = 1 To sys_cnt
      sys_col = sys_num + sys_i - 1
      If Cells(qcode_cnt, sys_col) <> "" Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",System(RC5:RC7)" & ",���݂̃o�[�W�����ł�E~G��͎g�p���Ȃ��ł��������B" & vbCrLf
      End If
    Next sys_i

    '�������w��ݖ�ƊY�������ݖ�̃Z���N�g�󋵂̓���`�F�b�N
    If Cells(qcode_cnt, rcode_col).Value <> "" Then
      Set rcode_cell = ws_setup.Range(Cells(setup_row, setup_col), Cells(max_row, setup_col)). _
      Find(What:=Cells(qcode_cnt, rcode_col).Value)
      If rcode_cell Is Nothing Then
        err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",����(RC2)" & ",�����w�肳�ꂽQCODE��������܂���B" & vbCrLf
      Else
        rcode_row = rcode_cell.Row
        If Cells(qcode_cnt, selcode_col_1).Value <> Cells(rcode_row, selcode_col_1).Value Or _
          Cells(qcode_cnt, selval_col_1).Value <> Cells(rcode_row, selval_col_1).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����@(RC10:RC11)" & ",�����@�i�Z���N�g�j�̎w�肪�Y���̎����ݖ��QCODE�m" _
          & Cells(rcode_row, setup_col).Value & "�n�ƈ�v���܂���B" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_2).Value <> Cells(rcode_row, selcode_col_2).Value Or _
          Cells(qcode_cnt, selval_col_2).Value <> Cells(rcode_row, selval_col_2).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����A(RC12:RC13)" & ",�����A�i�Z���N�g�j�̎w�肪�Y���̎����ݖ��QCODE�m" _
          & Cells(rcode_row, setup_col).Value & "�n�ƈ�v���܂���B" & vbCrLf
        End If
        If Cells(qcode_cnt, selcode_col_3).Value <> Cells(rcode_row, selcode_col_3).Value Or _
          Cells(qcode_cnt, selval_col_3).Value <> Cells(rcode_row, selval_col_3).Value Then
          err_msg = err_msg & Cells(qcode_cnt, setup_col).Value & "," & Format(qcode_cnt) & "�s��" & ",�����B(RC14:RC15)" & ",�����B�i�Z���N�g�j�̎w�肪�Y���̎����ݖ��QCODE�m" _
          & Cells(rcode_row, setup_col).Value & "�n�ƈ�v���܂���B" & vbCrLf
        End If
      End If
    End If
  End If
Next qcode_cnt

Cells(setup_row, setup_col).Select

Application.ScreenUpdating = True
preset_gcode = ws_mainmenu.Cells(gcode_row, gcode_col).Value
If err_msg <> "�y�ݒ��ʂ̃G���[���X�g�z" & vbCrLf & "QCODE,�s��,�G���[����,�G���[���e" & vbCrLf Then
  Application.DisplayAlerts = False
  ' �G���[���b�Z�[�W�o�̓t�@�C���̍쐬
  Workbooks.Add
  Set wb_preset = ActiveWorkbook
  Set ws_preset = wb_preset.Worksheets("Sheet1")
  Columns("A").NumberFormat = "@"
  If Dir(file_path & "\4_LOG\" & preset_gcode & "_�ݒ���NG.xlsx") = "" Then
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_�ݒ���NG.xlsx"
  Else
    Kill file_path & "\4_LOG\" & preset_gcode & "_�ݒ���NG.xlsx"
    wb_preset.SaveAs Filename:=file_path & "\4_LOG\" & preset_gcode & "_�ݒ���NG.xlsx"
  End If
  preset_row = Split(err_msg, vbCrLf)
  For preset_i = 0 To UBound(preset_row)
    preset_col = Split(preset_row(preset_i), ",")
    For preset_j = 0 To UBound(preset_col)
      ws_preset.Cells(preset_i + 1, preset_j + 1).Value = preset_col(preset_j)
    Next preset_j
  Next preset_i
  ws_preset.Activate
  ws_preset.Cells.Select
  With Selection.Font
    .Name = "Takao�S�V�b�N"
    .Size = 11
  End With
  Columns("B:D").EntireColumn.AutoFit
  ws_preset.Cells(1, 1).Select
  wb_preset.Save
  MsgBox "�ݒ��ʂ̓��e�ɃG���[������܂��B" & vbCrLf & _
  file_path & "\4_LOG\" & preset_gcode & "_�ݒ���NG.xlsx ���m�F���Ă��������B", vbExclamation, "MCS 2020 - Setup_Check"
  Application.DisplayAlerts = True
  ws_preset.Activate
  Application.StatusBar = False
End
  End If
  Application.StatusBar = False
  Set ws_preset = Nothing
  Set wb_preset = Nothing
End Sub

Sub Filepath_Get()
  '--------------------------------------------------------------------------------------------------'
  '�@�t�@�C���p�X�̎擾�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.10�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2017.05.18�@'
  '--------------------------------------------------------------------------------------------------'
  Application.StatusBar = "�t�@�C���p�X�擾�E������..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
    "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
    "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Input As #1
    Line Input #1, file_path
    Close #1
    ope_code = Cells(gcode_row, gcode_col)
  Else
    MsgBox "�ݒ�t�@�C���m" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini�n��������܂���B" _
    & vbCrLf & "�����ݒ��������x�s���Ă��������B", vbExclamation, "MCS 2020 - Filepath_Get"
    End
  End If
  Application.ScreenUpdating = True
  Application.StatusBar = False
End Sub

Sub Indata_Open()
  Dim indata_fn, revdata_fn As String
  '--------------------------------------------------------------------------------------------------'
  '�@���̓f�[�^�̃I�[�v��        �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.11�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.04.08�@'
  '--------------------------------------------------------------------------------------------------'
  '�y�T�v�z�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '�@���̃v���V�[�W���ŃI�[�v�������t�@�C���́A�ȉ��̃p�^�[���̃t�@�C���̂݁@�@�@�@�@�@�@�@�@�@�@�@'
  '�@�E���̓f�[�^�t�@�C�� ...... *IN.xlsx�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '�@�E�C����f�[�^�t�@�C�� .... *RE.xlsx                  �@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '�@�ECall���F���̓f�[�^�̏C���mModule04�n                                                          '
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "���̓f�[�^ �I�[�v����..."
  '    Application.ScreenUpdating = False    ' �t�@�C���I�[�v���̐i��������ꂽ�����悢�̂ŃR�����g�A�E�g

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & Cells(gcode_row, gcode_col) & "RE.xlsx") <> "" Then
    revdata_fn = Cells(gcode_row, gcode_col) & "RE.xlsx"
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & revdata_fn
    Set wb_indata = Workbooks(revdata_fn)
    Set ws_indata = wb_indata.Worksheets("Sheet1")
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
  ElseIf Dir(Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & Cells(gcode_row, gcode_col) & "IN.xlsx") <> "" Then
    indata_fn = Cells(gcode_row, gcode_col) & "IN.xlsx"
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & indata_fn
    Set wb_indata = Workbooks(indata_fn)
    Set ws_indata = wb_indata.Worksheets("Sheet1")
    Set wb_outdata = Nothing
    Set ws_outdata = Nothing
  Else
    MsgBox "���̓f�[�^�t�@�C�������݂��܂���B", vbExclamation, "MCS 2020 - Indata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Outdata_Open()
  '--------------------------------------------------------------------------------------------------'
  '�@���H��f�[�^�̃I�[�v���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.14�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.04.08�@'
  '--------------------------------------------------------------------------------------------------'
  '�y�T�v�z�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '�@���̃v���V�[�W���ŃI�[�v�������t�@�C���́A�W�v�ݒ�t�@�C���Ŏw�肳��Ă���t�@�C���@�@�@�@�@�@'
  '�@Call���F�W�v�T�}���[�f�[�^�̍쐬�mModule52�n�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "���H��f�[�^ �I�[�v����..."
  '    Application.ScreenUpdating = False

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn) <> "" Then
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn
    Set wb_outdata = ActiveWorkbook
    Set ws_outdata = wb_outdata.Worksheets("Sheet1")
    Set wb_indata = Nothing
    Set ws_indata = Nothing
  Else
    MsgBox "���H��f�[�^�t�@�C���m" & outdata_fn & "�n�����݂��܂���B", vbExclamation, "MCS 2020 - Outdata_Open"
  End
End If

Call Datacol_Get

'    Application.ScreenUpdating = True
Application.StatusBar = False
End Sub

Sub Datafile_Open()
  '--------------------------------------------------------------------------------------------------'
  '�@�f�[�^�t�@�C���̃I�[�v���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.27�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.04.08�@'
  '--------------------------------------------------------------------------------------------------'
  '�y�T�v�z�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '�@���̃v���V�[�W���ŃI�[�v�������t�@�C���́A���[�U�[������͂��ꂽ�t�@�C���@�@�@�@�@�@�@�@�@�@�@'
  '�@Call���F���̓f�[�^�̃��W�b�N�`�F�b�N�mModule05�n�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = status_msg

  wb.Activate
  ws_mainmenu.Select

  If Dir(Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn) <> "" Then
    Workbooks.Open Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA\" & outdata_fn
    Set wb_outdata = ActiveWorkbook
    Set ws_outdata = wb_outdata.Worksheets("Sheet1")
    Set wb_indata = Nothing
    Set ws_indata = Nothing
  Else
    MsgBox "�w�肵���t�@�C���m" & outdata_fn & "�n�����݂��܂���B", vbExclamation, "MCS 2020 - Datafile_Open"
  End
End If

Call Datacol_Get

Application.StatusBar = False
End Sub

Sub Datacol_Get()
  Dim wb_data As Workbook
  Dim ws_data As Worksheet

  Dim max_row, max_col As Long
  Dim i_cnt, setup_cnt As Long
  Dim match_flg As Long
  Dim pass_flg As Integer
  '--------------------------------------------------------------------------------------------------'
  '�@�f�[�^�t�@�C���̗�ԍ����擾�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.27�@'
  '�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2018.04.04�@'
  '--------------------------------------------------------------------------------------------------'
  Application.ScreenUpdating = False

  If Not wb_indata Is Nothing Then
    Set wb_data = wb_indata
    Set ws_data = ws_indata
  ElseIf Not wb_outdata Is Nothing Then
    Set wb_data = wb_outdata
    Set ws_data = ws_outdata
  End If

  wb_data.Activate
  ws_data.Select

  max_col = ws_data.Cells(1, Columns.Count).End(xlToLeft).Column
  ReDim incol_array(max_col, 2)

  ' �f�[�^�t�@�C���̗�ԍ����擾
  For i_cnt = 1 To max_col
    incol_array(i_cnt, 1) = ws_data.Cells(1, i_cnt)
    incol_array(i_cnt, 2) = i_cnt
  Next i_cnt

  wb.Activate
  ws_setup.Select

  max_row = ws_setup.Cells(Rows.Count, setup_col).End(xlUp).Row
  pass_flg = 0

  Range(Cells(setup_row, 3), Cells(max_row, 3)).ClearContents

  For setup_cnt = setup_row To max_row
    match_flg = 0
    If Mid(ws_setup.Cells(setup_cnt, setup_col), 1, 1) <> "*" Then
      For i_cnt = 1 To max_col
        If ws_setup.Cells(setup_cnt, setup_col) = incol_array(i_cnt, 1) Then
          ws_setup.Cells(setup_cnt, incol_col) = incol_array(i_cnt, 2)
          match_flg = 1
          Exit For
        End If
      Next i_cnt
      If pass_flg = 0 Then
        If match_flg = 0 Then
          MsgBox "QCODE�m" & ws_setup.Cells(setup_cnt, 1) & "�n���A�f�[�^�t�@�C���ɂ���܂���B", vbExclamation, "MCS 2020 - Datacol_Get"
          wb.Activate
          ws_mainmenu.Select
          Application.StatusBar = False
        End
      End If
    End If
  ElseIf ws_setup.Cells(setup_cnt, setup_col) = "*���H��" Then
    pass_flg = 1
    ws_setup.Cells(setup_cnt, incol_col) = incol_array(i_cnt, 2) + 1
  End If
Next setup_cnt

Application.ScreenUpdating = True
End Sub

Sub Setup_Hold()
  Dim max_row, max_col As Long

  '    Dim q_data() As question_data   ' QCODE���̃f�[�^��S�Ď擾
  Dim work_count As Long          ' �����񐔃J�E���g�p�ϐ�
  Dim work_string As String       ' ��Ɨp������ϐ�
  Dim ctwork_count As Long        ' �ݖ⎈�i�[�p�J�E���g�ϐ�
  '    Dim writing_target As Long      ' �z�񏑂����݈ʒu�i�[�p�ϐ�

  '--------------------------------------------------------------------------------------------------'
  '�@�ݒ��ʂ̏����z�[���h�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
  '--------------------------------------------------------------------------------------------------'
  '�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.11�@'
  '�@�ŏI�ҏW�ҁ@���R�@���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2017.06.26�@'
  '--------------------------------------------------------------------------------------------------'
  Call Starting_Mcs2017
  Application.StatusBar = "�ݒ��ʏ�� �z�[���h��..."
  Application.ScreenUpdating = False

  wb.Activate
  ws_setup.Select
  '
  '    max_row = Cells(Rows.Count, setup_col).End(xlUp).Row
  '    max_col = Cells(1, Columns.Count).End(xlToLeft).Column
  '
  '    setup_array = Range(Cells(1, 1), Cells(max_row, max_col + 1))
  '

  ' --- 20170414 ���R�� �ǉ����� ---------------------------------------------------------------------'

  ' QCODE�̐����擾
  max_row = ws_setup.Cells(Rows.Count, setup_col).End(xlUp).Row

  ' QCODE���ɍ��킹�ăf�[�^�z�񐔂��Ē�`
  ReDim q_data(max_row)

  ' 20170419 QCODE���ɍ��킹�Č����p�f�[�^�z�񐔂��Ē�`
  ReDim str_code(max_row)

  ' �z�񏑂����݈ʒu��������
  'writing_target = 0

  ' QCODE�̏���S�Ď擾����
  For work_count = setup_row To max_row

    ' �z�񏑂����݈ʒu���C���N�������g
    'writing_target = writing_target + 1

    ' 20170419 �����p�z���QCODE���i�[
    str_code(work_count) = ws_setup.Cells(work_count, 1).Value

    q_data(work_count).q_code = ws_setup.Cells(work_count, 1).Value                 ' QCODE
    q_data(work_count).r_code = ws_setup.Cells(work_count, 2).Value                 ' ����
    q_data(work_count).m_code = ws_setup.Cells(work_count, 8).Value                 ' MCODE

    work_string = Trim(ws_setup.Cells(work_count, 9).Value)                         ' �`�����擾

    ' ��ԁA���ԁA�N���A
    If Mid(work_string, 2, 1) = "M" Or _
      Mid(work_string, 2, 1) = "A" Or _
      Mid(work_string, 2, 1) = "C" Then

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT��
      q_data(work_count).ct_loop = Val(Mid(work_string, 3))                   ' LCT��
      q_data(work_count).q_format = Mid(work_string, 1, 2)                    ' �`��

      ' L�̂�
    Else

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT��
      q_data(work_count).ct_loop = Val(Mid(work_string, 2))                   ' LCT��
      q_data(work_count).q_format = Mid(work_string, 1, 1)                    ' �`��

    End If

    ' �`���ɉ����ď�����ύX
    Select Case q_data(work_count).q_format

      ' �T���v���i���o�[��Ǘ��R�[�h��
    Case "C"

      ' �V���O���A���T�[
    Case "S"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT��

      ' �}���`�A���T�[
    Case "M"

      q_data(work_count).ct_count = ws_setup.Cells(work_count, 16).Value      ' CT��
      q_data(work_count).ct_loop = ws_setup.Cells(work_count, 16).Value       ' LCT��


      ' ���A���A���T�[
    Case "R"

      q_data(work_count).r_byte = Val(Mid(work_string, 2))                    ' ����BYTE��

      ' �t���[�A���T�[
    Case "F"

      ' �g�J�[�\��
    Case "H"

      q_data(work_count).r_byte = 3                                           ' ����BYTE��

      ' �I�[�v���A���T�[
    Case "O"

      ' LM���܂ߏ������s��Ȃ�����
    Case Else

  End Select

  q_data(work_count).r_unit = ws_setup.Cells(work_count, 4).Value                 ' �����P�ʁi2017.05.10 �ǉ��j

  q_data(work_count).sel_code1 = ws_setup.Cells(work_count, 10).Value             ' �Z���N�g�@ QCODE

  ' sel_code1�����͂���Ă�����
  If q_data(work_count).sel_code1 <> "" Then
    q_data(work_count).sel_value1 = Val(ws_setup.Cells(work_count, 11).Value)   ' �Z���N�g�@ VALUE
  End If

  q_data(work_count).sel_code2 = ws_setup.Cells(work_count, 12).Value             ' �Z���N�g�A QCODE

  ' sel_code2�����͂���Ă�����
  If q_data(work_count).sel_code2 <> "" Then
    q_data(work_count).sel_value2 = Val(ws_setup.Cells(work_count, 13).Value)   ' �Z���N�g�A VALUE
  End If

  q_data(work_count).sel_code3 = ws_setup.Cells(work_count, 14).Value             ' �Z���N�g�B QCODE

  ' sel_code3�����͂���Ă�����
  If q_data(work_count).sel_code3 <> "" Then
    q_data(work_count).sel_value3 = Val(ws_setup.Cells(work_count, 15).Value)   ' �Z���N�g�B VALUE
  End If

  ' �O�J�e�S���[���܂ސݖ₩�𔻒�
  '        If Trim(ws_setup.Cells(work_count, 17).Value) <> "" Then
  '            q_data(work_count).ct_0flg = True                                           ' 0CT�t���O
  '        End If

  q_data(work_count).q_title = ws_setup.Cells(work_count, 18).Value               ' �\��

  ' �ݖ⎈�����݂��鎞
  If q_data(work_count).ct_count <> 0 Then

    ' �J�e�S���[�����ݖ⎈���擾����
    For ctwork_count = 1 To q_data(work_count).ct_count

      q_data(work_count).q_ct(ctwork_count) = _
      ws_setup.Cells(work_count, 18 + ctwork_count).Value                     ' �ݖ⎈

    Next ctwork_count

  End If

  ' ���̓f�[�^��ԍ����L������Ă��鎞
  If ws_setup.Cells(work_count, 3).Value <> "" Then
    q_data(work_count).data_column = Val(ws_setup.Cells(work_count, 3).Value)   ' ���̓f�[�^��ԍ�
  End If

Next work_count

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2017.04.19  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2017.05.02�@'
' QCODE�}�b�`���O�p�t�@���N�V����                                                                  '
' ���@�� String�^  QCODE                                                                           '
' �߂�l Integer�^ ��ԍ�                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_Match(ByVal match_code As String) As Integer

  ' �G���[�𖳎����ď������s��
  On Error Resume Next

  ' QCODE����������ԍ���Ԃ�
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1

  ' QCODE���}�b�`���O���Ȃ��ꍇ�̓G���[�\��
  If Qcode_Match = 0 Then
    MsgBox "QCODE�m" & match_code & "�n�͐ݒ��ʂɑ��݂��܂���B" & _
    vbCrLf & "�v���O�����������I�����܂��B", vbCritical, "MCS 2020 - Qcode_Match�y�����I���z"
    wb.Activate
    ws_mainmenu.Select

    '2018/05/15 - �ǋL ==========================
    Dim myWB As Workbook
    If Workbooks.Count > 1 Then
      Application.DisplayAlerts = False
      Application.Visible = True
      For Each myWB In Workbooks
        If myWB.Name <> ActiveWorkbook.Name Then
          myWB.Close
        End If
      Next
      Application.DisplayAlerts = True
    End If
    '============================================

    '2018/06/13 - �ǋL ==========================
    ' SUM�t�H���_����0�o�C�g�t�@�C�����폜
    Dim sum_file As String
    sum_file = Dir(file_path & "\SUM\*_sum.xlsx")
    Do Until sum_file = ""
      DoEvents
      If FileLen(file_path & "\SUM\" & sum_file) = 0 Then
        Kill file_path & "\SUM\" & sum_file
      End If
      sum_file = Dir()
    Loop
    '============================================
    Call Finishing_Mcs2017
  End
End If

End Function

'--------------------------------------------------------------------------------------------------'
' ��@���@�� ���R ��                                                           �쐬��  2018.06.25  '
' �ŏI�ҏW�� ���R ���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@   �@�ҏW���@2018.06.25�@'
' QCODE�}�b�`���O�p�t�@���N�V�����i�G���[�߂��o�[�W�����j                                          '
' ���@�� String�^  QCODE                                                                           '
' �߂�l Integer�^ ��ԍ�                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Qcode_MatchE(ByVal match_code As String) As Integer
  ' �G���[�𖳎����ď������s��
  On Error Resume Next
  ' QCODE����������ԍ���Ԃ�
  Qcode_Match = Application.WorksheetFunction.Match(match_code, str_code, 0) - 1
End Function

'--------------------------------------------------------------------------------------------------'
' �쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@ �쐬���@2017.06.20�@'
' �ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@ �ҏW���@2017.06.20�@'
' �X�e�[�^�X�o�[�h�b�g�A�N�V�����p�t�@���N�V����                                                   '
' ���@�� Integer�^ �J�E���^�[                                                                      '
' �߂�l String�^  �h�b�g                                                                          '
'--------------------------------------------------------------------------------------------------'
Public Function Status_Dot(ByVal i_cnt As Integer) As String
  If i_cnt Mod 4 = 0 Then
    Status_Dot = ""
  Else
    Status_Dot = String(i_cnt Mod 4, ".")
  End If
End Function

