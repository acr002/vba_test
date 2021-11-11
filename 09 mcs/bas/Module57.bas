Attribute VB_Name = "Module57"
Option Explicit
    Dim wb_tabinst As Workbook
    Dim ws_tabinst As Worksheet
    Dim tabinst_fn As String
    Dim face_qcode As String
    Dim period_pos As Long

    Dim t_seq As Long    '�\��
    Dim t_r As Long      '�\���s�J�E���g
    Dim t_ra As Long     '�����w��s�J�E���g
    Dim ct_f As Integer  '�����J�e�S���C�Y�t���O

    Dim t_crs As String  '��3����QCODE
    Dim t_cnt As Long    '��3���̃J�e�S���[��

Sub Triplecross_Setting()
    Dim d_index As Long
    Dim r_code As Integer
    Dim i_cnt As Long
'--------------------------------------------------------------------------------------------------'
'�@�R�d�N���X�p�W�v�ݒ�t�@�C���̍쐬�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2019.10.02�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check

    wb.Activate
    ws_mainmenu.Select
step00:
    tabinst_fn = InputBox("�쐬����W�v�ݒ�t�@�C���̃t�@�C������" & vbCrLf & "���͂��Ă��������B" & vbCrLf & vbCrLf & "�y��zA01030C01.xlsx �Ȃ�", "MCS 2020 - 3�d�N���X�p�W�v�ݒ�t�@�C���̍쐬")
    If tabinst_fn = "" Then
        Application.StatusBar = False
        End
    End If

    period_pos = InStrRev(tabinst_fn, ".")
    If period_pos > 0 Then
        If LCase(Mid(tabinst_fn, period_pos + 1)) <> "xlsx" Then
            MsgBox "�t�@�C���`��������������܂���B" _
             & vbCrLf & "�g���q�� xlsx ���w�肵�Ă��������B", vbExclamation, "MCS 2020 - Tabulation_Setting"
            GoTo step00
        End If
    Else
        tabinst_fn = tabinst_fn & ".xlsx"
    End If
    
    Call Setup_Hold
'2019.10.9 - �ǉ�����---------------------------------------------------
    ws_setup.Select
    face_qcode = ""
    face_qcode = InputBox("�\���ɐݒ肷��QCODE����͂��Ă��������B" & vbCrLf & vbCrLf & "�y��zF01 �Ȃ�" & vbCrLf & "�������͂̏ꍇ�́A�P���W�v�̏W�v�ݒ�ƂȂ�܂��B", "MCS 2020 - 3�d�N���X�p�W�v�ݒ�t�@�C���̍쐬")
    
    If StrPtr(face_qcode) = 0 Then  '�L�����Z������
        wb.Activate
        ws_mainmenu.Select
        Application.StatusBar = False
        End
    End If
    
    If face_qcode <> "" Then
        d_index = Qcode_Match(face_qcode)
        If (q_data(d_index).q_format = "R") Or (q_data(d_index).q_format = "H") _
         Or (q_data(d_index).q_format = "C") Or (q_data(d_index).q_format = "F") _
         Or (q_data(d_index).q_format = "O") Then
            face_qcode = ""
        End If
    End If
'-----------------------------------------------------------------------
    
step10:
    ws_setup.Select
    t_crs = ""
    t_crs = InputBox("��3���i�W�v�������j�ɐݒ肷��QCODE����͂��Ă��������B" & vbCrLf & vbCrLf & "�y��zKBN �Ȃ�", "MCS 2020 - 3�d�N���X�p�W�v�ݒ�t�@�C���̍쐬")
    
    If StrPtr(t_crs) = 0 Then  '�L�����Z������
        wb.Activate
        ws_mainmenu.Select
        Application.StatusBar = False
        End
    End If
    
    If t_crs <> "" Then
        r_code = MsgBox("�W�v�ݒ�t�@�C�����3���̃J�e�S���[���Ƃɕ�������" & vbCrLf & "�쐬���܂����H" _
         & vbCrLf & vbCrLf & "�u�͂��v�@�� �W�v�ݒ�t�@�C���𕪊��쐬" & vbCrLf & "�u�������v�� �W�v�ݒ�t�@�C����1�t�@�C���Ƃ��č쐬", _
         vbYesNoCancel + vbQuestion, "MCS 2020 - Spreadsheet_Creation")
        
        d_index = Qcode_Match(t_crs)
        If (q_data(d_index).q_format = "R") Or (q_data(d_index).q_format = "H") _
         Or (q_data(d_index).q_format = "C") Or (q_data(d_index).q_format = "F") _
         Or (q_data(d_index).q_format = "O") Then
            MsgBox "��3���̌`�����m�F���Ă��������B" & vbCrLf & vbCrLf & _
             "�yTIPS�z" & vbCrLf & "��3���ɐݒ�ł���QCODE�̌`���́A" & vbCrLf & "�mSA�n�mMA�n�mLMA�n�ƂȂ��Ă���܂��B", vbInformation, "MCS 2020 - Triplecross_Setting"
            GoTo step10
        Else
            t_cnt = q_data(d_index).ct_count
        End If
    Else
        MsgBox "��3���̎w�肪����܂���B" & vbCrLf & vbCrLf & _
         "�yTIPS�z" & vbCrLf & "�ʏ�̒P���W�v�\�E�N���X�W�v�\���쐬����ꍇ�́A" & vbCrLf & "�u�W�v�ݒ�t�@�C���̍쐬�v����쐬���Ă��������B", vbInformation, "MCS 2020 - Triplecross_Setting"
        GoTo step10
    End If
    
    Application.StatusBar = "3�d�N���X�p�W�v�ݒ�t�@�C�� �쐬��..."
    Application.ScreenUpdating = False

    wb.Activate
    ws_mainmenu.Select

    If r_code = vbYes Then
        Call split_tabinst    ' �����쐬������
        GoTo step99
    End If

' �W�v�ݒ�t�@�C��1�쐬����
    Workbooks.Add
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & tabinst_fn
    Application.DisplayAlerts = True

    If Err.Number <> 0 Then
        ActiveWorkbook.Close
        Open file_path & "\3_FD\" & tabinst_fn For Append As #1
        Close #1
        If Err.Number = 70 Then
            MsgBox tabinst_fn & " �́A���łɊJ����Ă��܂��B" _
            & vbCrLf & "�t�@�C������Ă���Ď��s���Ă��������B", vbExclamation, "MCS 2020 - Tabulation_Setting"
        End If
        End
    End If

    Set wb_tabinst = ActiveWorkbook
    Set ws_tabinst = wb_tabinst.ActiveSheet

    Call Inst_Header

' ��������W�v�ݒ�t�@�C���쐬�̃R�[�f�B���O(�L��֥`)
    t_seq = 0
    For i_cnt = 1 To t_cnt
        For t_r = 3 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
            If (ws_setup.Cells(t_r, 1).Value <> "weight") And _
             (Mid(ws_setup.Cells(t_r, 1).Value, 1, 2) <> "SE") Then ' QCODE�mweight�n�ƁmSE�n�͂��܂�i���H��Z���N�g�j�͏W�v�ݒ�t�@�C���ɏo�͂��Ȃ�
                If Left(ws_setup.Cells(t_r, 1).Value, 1) <> "*" Then ' QCODE��̐擪�A�X�^���X�N�s�͏������Ȃ�
                    Select Case Left(ws_setup.Cells(t_r, 9).Value, 1)
                        Case "S", "M", "L"  ' SA,MA,LMA��QCODE��\���ɏo��
                            If ws_setup.Cells(t_r, 2).Value = "" Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                            Else
                                For t_ra = t_r + 1 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                    If ws_setup.Cells(t_r, 2).Value = ws_setup.Cells(t_ra, 1).Value Then
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                                    End If
                                Next t_ra
                            End If
                            
                            ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                            ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                            
                            '2018/9/13 - �ǉ����ڂ̂��߂̏���
                            d_index = Qcode_Match(ws_setup.Cells(t_r, 1).Value)
                            If Left(q_data(d_index).q_format, 1) = "S" Then
                                If q_data(d_index).ct_count <= 5 Then
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' �~�O���t
                                Else
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                End If
                            Else
                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                            End If
                        Case "R", "H"  ' RA,HC�̓J�e�S���C�Y���QCODE�����[�v�ŒT���ĕ\���ɏo��
                            ct_f = 0
                            For t_ra = t_r To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value Then
                                    If Left(ws_setup.Cells(t_ra, 1).Value, 1) <> "*" Then
                                        ct_f = 1
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_ra, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_ra, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                                        
                                        ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                                        ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                                        
                                        ' 2018/9/13 - �ǉ����ڂ̂��߂̏���
                                        d_index = Qcode_Match(ws_setup.Cells(t_ra, 1).Value)
                                        If Left(q_data(d_index).q_format, 1) = "S" Then
                                            If q_data(d_index).ct_count <= 5 Then
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' �~�O���t
                                            Else
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                            End If
                                        Else
                                            ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                        End If
                                    End If
                                End If
                            Next t_ra
                            If ct_f = 0 Then
                                For t_ra = 3 To t_r - 1
                                    If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value And t_r >= 3 Then
                                        ct_f = 1
                                    End If
                                Next t_ra
                            End If
                            If ct_f = 0 Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 1).Value
                                ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                            End If
                        Case "C", "F", "O"  ' CODE,FA,OA�͏W�v�ݒ�t�@�C���ɏo�͂��Ȃ�
                        Case Else
                    End Select
                End If
            End If
        Next t_r
    Next i_cnt

    wb_tabinst.Activate
    ws_tabinst.Select
    ws_tabinst.Range(Cells(7, 1), Cells(t_seq + 6, 25)).Select
    With Selection
        .Borders.LineStyle = xlContinuous
    End With

    ' �t�@�C���㕔�̏����ƑS�̓I�ȏ���
    ws_tabinst.Cells(2, 1).Value = "�W�v����f�[�^�t�@�C����"

    Range("A2:C2").Select
    Selection.Merge

    Range("A2:G2").Select
    With Selection.Font
        .Color = 16724787
        .TintAndShade = 0
    End With
    
    ws_tabinst.Cells(2, 9).Value = "�y�E�G�C�g�W�v�̐ݒ�z"
    ws_tabinst.Cells(2, 10).Value = "�Ȃ�"
    ws_tabinst.Cells(2, 11).Value = "����"

    Range("L2").Select
    ws_tabinst.Cells(2, 12).Value = "�Ȃ�"
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    With Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$J$2:$K$2"
    End With
    
    Range("A2:C2").Select
    Selection.Copy
    Range("I2:K2").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

    Range("E:Y").ColumnWidth = 7.13
    Range("E:Y").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    Rows(1).RowHeight = 6.75
    Rows(2).RowHeight = 28.5
    Rows(3).RowHeight = 6.75
        
    Range("D2:G2").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws_tabinst.Cells(2, 4) = ws_mainmenu.Cells(3, 8) & "OT.xlsx"
    
    ActiveWindow.DisplayGridlines = False
    ws_tabinst.Cells(1, 1).Select

' �㏑���ۑ����ăt�@�C�������
    wb_tabinst.Activate
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

    Set wb_tabinst = Nothing
    Set ws_tabinst = Nothing
    
' �V�X�e�����O�̏o��
step99:
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "22"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 22"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �W�v�ݒ�t�@�C���̍쐬�F�쐬�t�@�C���m" & tabinst_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "�W�v�ݒ�t�@�C�� " & tabinst_fn & " �̍쐬���������܂����B", vbInformation, "MCS 2020 - Triplecross_Setting"
End Sub

Private Sub Inst_Header()
' �w�b�_�[�̍쐬
    Cells.Select
    Selection.NumberFormatLocal = "@"
    With Selection.Font
        .Name = "Takao�S�V�b�N"
        .Size = 11
    End With
    
    Range("A4:Y6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

    Range("A4:A6").Select
    Selection.Merge

    Range("B4:B5").Select
    Selection.Merge

    Range("C4:C5").Select
    Selection.Merge

    Range("D4:D5").Select
    Selection.Merge

    Range("E4:E5").Select
    Selection.Merge

    Range("F4:F5").Select
    Selection.Merge

    Range("G4:G5").Select
    Selection.Merge

    Range("A4:D6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

    Range("E4:G6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13421823
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 204
        .TintAndShade = 0
    End With

    Range("H4:P4").Select
    Selection.Merge
    Range("H4:P6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16772300
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -52429
        .TintAndShade = 0
    End With

    Range("Q4:R4").Select
    Selection.Merge
    Range("Q4:R6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434848
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 32768
        .TintAndShade = 0
    End With

    Range("S4:U4").Select
    Selection.Merge
    Range("S4:U6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13434879
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = 13158
        .TintAndShade = 0
    End With
    
    Range("V4:V6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10079487
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -16764007
        .TintAndShade = 0
    End With

    Range("W4:X4").Select
    Selection.Merge
    Range("W4:X6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764006
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Color = -6737152
        .TintAndShade = 0
    End With

    Range("Y4:Y4").Select
    Selection.Merge
    Range("Y4:Y6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777215
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With

    Range("A4:Y6").Borders.LineStyle = xlContinuous
    Range("A6:Y6").Borders(xlEdgeTop).LineStyle = xlLineStyleNone

    Range("A4") = "�\��"
    Range("B4") = "�\��"
    Range("B6") = "(QCODE)"
    Range("C4") = "�\��"
    Range("C6") = "(QCODE)"
    Range("D4") = "����"
    Range("D6") = "(QCODE)"
    Range("E4") = "�\���\��"
    Range("E6") = "(N/E)"
    Range("F4") = "�\��NA"
    Range("F6") = "(N)"
    Range("G4") = "�ꐔ"
    Range("G6") = "(Y)"
    Range("H4") = "�����ݖ�o�͎w��"
    Range("H5") = "���v"
    Range("H6") = "(Y)"
    Range("I5") = "����"
    Range("I6") = "(Num/Y)"
    Range("J5") = "�W���΍�"
    Range("J6") = "(Num/Y)"
    Range("K5") = "�ŏ��l"
    Range("K6") = "(Y)"
    Range("L5") = "��P�l����"
    Range("L6") = "(Y)"
    Range("M5") = "�����l"
    Range("M6") = "(Y)"
    Range("N5") = "��R�l����"
    Range("N6") = "(Y)"
    Range("O5") = "�ő�l"
    Range("O6") = "(Y)"
    Range("P5") = "�ŕp�l"
    Range("P6") = "(Y)"
    Range("Q4") = "����"
    Range("Q5") = "QCODE"
    Range("Q6") = "(QCODE)"
    Range("R5") = "�l"
    Range("R6") = "(Num)"
    Range("S4") = "�\���I�v�V����"
    Range("S5") = "������"
    Range("S6") = "(Y)"
    Range("T5") = "�L����"
    Range("T6") = "(Y)"
    Range("U5") = "�q�׉�"
    Range("U6") = "(Y)"

'2018/9/13 - �ǉ�����
    Range("V4") = "TOP1"
    Range("V5") = "�}�[�L���O"
    Range("V6") = "(Y/A)"
    Range("W4") = "GT�\�[�g"
    Range("W5") = "�~��"
    Range("W6") = "(Y)"
    Range("X5") = "���OCT"
    Range("X6") = "(Num)"
    Range("Y4") = "�O���t"
    Range("Y5") = "���"
    Range("Y6") = "(Num)"

    Range("A7").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub split_tabinst()
' �W�v�ݒ�t�@�C�������쐬���� - 2019.10.2 �ǋL
    Dim fd As String
    Dim d_index As Long
    Dim r_code As Integer
    Dim i_cnt As Long
    
    For i_cnt = 1 To t_cnt
        Workbooks.Add
        Application.DisplayAlerts = False
        
        fd = Dir(file_path & "\3_FD\CRS", vbDirectory)
        If fd = "" Then
            MkDir file_path & "\3_FD\CRS"
        End If
        
        ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\CRS\" & Format(i_cnt, "00") & "_" & tabinst_fn
        Application.DisplayAlerts = True

        If Err.Number <> 0 Then
            ActiveWorkbook.Close
            Open file_path & "\3_FD\CRS\" & Format(i_cnt, "00") & "_" & tabinst_fn For Append As #1
            Close #1
            If Err.Number = 70 Then
                MsgBox Format(i_cnt, "00") & "_" & tabinst_fn & " �́A���łɊJ����Ă��܂��B" _
                 & vbCrLf & "�t�@�C������Ă���Ď��s���Ă��������B", vbExclamation, "MCS 2020 - Tabulation_Setting"
            End If
            End
        End If
        
        Set wb_tabinst = ActiveWorkbook
        Set ws_tabinst = wb_tabinst.ActiveSheet

        Call Inst_Header

        t_seq = 0
        For t_r = 3 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
            If (ws_setup.Cells(t_r, 1).Value <> "weight") And _
             (Mid(ws_setup.Cells(t_r, 1).Value, 1, 2) <> "SE") Then ' QCODE�mweight�n�ƁmSE�n�͂��܂�i���H��Z���N�g�j�͏W�v�ݒ�t�@�C���ɏo�͂��Ȃ�
                If Left(ws_setup.Cells(t_r, 1).Value, 1) <> "*" Then ' QCODE��̐擪�A�X�^���X�N�s�͏������Ȃ�
                    Select Case Left(ws_setup.Cells(t_r, 9).Value, 1)
                        Case "S", "M", "L"  ' SA,MA,LMA��QCODE��\���ɏo��
                            If ws_setup.Cells(t_r, 2).Value = "" Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                            Else
                                For t_ra = t_r + 1 To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                    If ws_setup.Cells(t_r, 2).Value = ws_setup.Cells(t_ra, 1).Value Then
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_r, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                                    End If
                                Next t_ra
                            End If
                            
                            ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                            ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                            
                            '2018/9/13 - �ǉ����ڂ̂��߂̏���
                            d_index = Qcode_Match(ws_setup.Cells(t_r, 1).Value)
                            If Left(q_data(d_index).q_format, 1) = "S" Then
                                If q_data(d_index).ct_count <= 5 Then
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' �~�O���t
                                Else
                                    ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                End If
                            Else
                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                            End If
                        Case "R", "H"  ' RA,HC�̓J�e�S���C�Y���QCODE�����[�v�ŒT���ĕ\���ɏo��
                            ct_f = 0
                            For t_ra = t_r To ws_setup.Cells(Rows.Count, 1).End(xlUp).Row
                                If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value Then
                                    If Left(ws_setup.Cells(t_ra, 1).Value, 1) <> "*" Then
                                        ct_f = 1
                                        t_seq = t_seq + 1
                                        ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                        ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                        ws_tabinst.Cells(t_seq + 6, 3).Value = ws_setup.Cells(t_ra, 1).Value
                                        ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_ra, 2).Value
                                        ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                        ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                                        
                                        ws_tabinst.Cells(t_seq + 6, 17).Value = t_crs
                                        ws_tabinst.Cells(t_seq + 6, 18).Value = i_cnt
                                        
                                        ' 2018/9/13 - �ǉ����ڂ̂��߂̏���
                                        d_index = Qcode_Match(ws_setup.Cells(t_ra, 1).Value)
                                        If Left(q_data(d_index).q_format, 1) = "S" Then
                                            If q_data(d_index).ct_count <= 5 Then
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "1"    ' �~�O���t
                                            Else
                                                ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                            End If
                                        Else
                                            ws_tabinst.Cells(t_seq + 6, 25).Value = "2"    ' ���O���t
                                        End If
                                    End If
                                End If
                            Next t_ra
                            If ct_f = 0 Then
                                For t_ra = 3 To t_r - 1
                                    If ws_setup.Cells(t_r, 1).Value = ws_setup.Cells(t_ra, 2).Value And t_r >= 3 Then
                                        ct_f = 1
                                    End If
                                Next t_ra
                            End If
                            If ct_f = 0 Then
                                t_seq = t_seq + 1
                                ws_tabinst.Cells(t_seq + 6, 1).Value = Format(t_seq, "0000")
                                ws_tabinst.Cells(t_seq + 6, 2).Value = face_qcode
                                ws_tabinst.Cells(t_seq + 6, 4).Value = ws_setup.Cells(t_r, 1).Value
                                ws_tabinst.Cells(t_seq + 6, 8).Value = "Y"
                                ws_tabinst.Cells(t_seq + 6, 9).Value = "1"    ' ����͏����_��P�ʂŏo��
                            End If
                        Case "C", "F", "O"  ' CODE,FA,OA�͏W�v�ݒ�t�@�C���ɏo�͂��Ȃ�
                        Case Else
                    End Select
                End If
            End If
        Next t_r
        wb_tabinst.Activate
        ws_tabinst.Select
        ws_tabinst.Range(Cells(7, 1), Cells(t_seq + 6, 25)).Select
        With Selection
            .Borders.LineStyle = xlContinuous
        End With

        ' �t�@�C���㕔�̏����ƑS�̓I�ȏ���
        ws_tabinst.Cells(2, 1).Value = "�W�v����f�[�^�t�@�C����"

        Range("A2:C2").Select
        Selection.Merge

        Range("A2:G2").Select
        With Selection.Font
            .Color = 16724787
            .TintAndShade = 0
        End With

        ws_tabinst.Cells(2, 9).Value = "�y�E�G�C�g�W�v�̐ݒ�z"
        ws_tabinst.Cells(2, 10).Value = "�Ȃ�"
        ws_tabinst.Cells(2, 11).Value = "����"

        Range("L2").Select
        ws_tabinst.Cells(2, 12).Value = "�Ȃ�"
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        With Selection.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=$J$2:$K$2"
        End With
    
        Range("A2:C2").Select
        Selection.Copy
        Range("I2:K2").Select
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        Range("E:Y").ColumnWidth = 7.13
        Range("E:Y").Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        Rows(1).RowHeight = 6.75
        Rows(2).RowHeight = 28.5
        Rows(3).RowHeight = 6.75
        
        Range("D2:G2").Select
        Selection.Merge
        With Selection
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        ws_tabinst.Cells(2, 4) = ws_mainmenu.Cells(3, 8) & "OT.xlsx"
    
        ActiveWindow.DisplayGridlines = False
        ws_tabinst.Cells(1, 1).Select
    
        ' �㏑���ۑ����ăt�@�C�������
        wb_tabinst.Activate
        Application.DisplayAlerts = False
        ActiveWorkbook.Save
        ActiveWorkbook.Close
        Application.DisplayAlerts = True

        Set wb_tabinst = Nothing
        Set ws_tabinst = Nothing
    
    Next i_cnt
End Sub
