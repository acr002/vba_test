Attribute VB_Name = "Module31"
Option Explicit
    Public ot_fn As String
    Public otpt_type As String
    Public ctz_type As Long, hdr_type As Long
    
    Dim otrd_wb As Workbook
    Dim otrd_ws As Worksheet, index_ws As Worksheet
    Dim rd_fn As String, newworkbook_fn As String, qcode As String
    Dim s_r As Long, oted_col  As Long, oted_row As Long, index_cnt As Long, fc_rw As Long
    Dim i As Long, j As Long, code_cnt As Long, ct_cnt As Long, ma_ed As Long, ct_no As Long, ct_len As Long
    Dim alrt_msg As String, ma_stamp As String, lbl_adr As String, ct_label As String, zero As String
    Dim adr_arr As Variant
    Dim fc_rng As Range

Public Sub RD_Creation()
'--------------------------------------------------------------------------------------------------'
'�@���[�f�[�^�̍쐬�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�c���@�`�W�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.05.17�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\1_DATA"
    
step00:
    wb.Activate
    ws_mainmenu.Select
    ot_fn = Application.GetOpenFilename("�f�[�^�t�@�C��,*.xlsx", , "�f�[�^�t�@�C�����J��")
    If ot_fn = "False" Then
        ' �L�����Z���{�^���̏���
        End
    ElseIf ot_fn = "" Then
        MsgBox "���[�f�[�^�t�@�C�����쐬����m�f�[�^�t�@�C���n��I�����Ă��������B", vbExclamation, "MCS 2020 - RD_Creation"
        Application.StatusBar = False
        GoTo step00
    End If
    
    Application.Visible = False
    Application.ScreenUpdating = False
    Load FrmRDsel
    FrmRDsel.StartUpPosition = 1
    FrmRDsel.Show
    FrmRDsel.Repaint
    Unload FrmRDsel
    
' �ΏۂƂȂ�t�@�C���́A��L�_�C�A���O����擾�ɕύX - 2018/06/01
    Open ot_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(Dir(ot_fn)).Close
    Else
        Workbooks.Open ot_fn
    End If
    
    rd_fn = "���[�f�[�^.xlsx"
    Open file_path & "\1_DATA\" & rd_fn For Append As #1
    Close #1
    If Err.Number > 0 Then
        Workbooks(rd_fn).Close
    End If
    
'�������烍�[�f�[�^�̍쐬�R�[�f�B���O(�L��֥`)
    ot_fn = Dir(ot_fn)
    Set otrd_wb = Workbooks(ot_fn)
    Set otrd_ws = otrd_wb.Worksheets(1)
    otrd_wb.Activate
    otrd_ws.Select
    otrd_ws.Name = "���[�f�[�^"
    oted_col = otrd_ws.Cells(1, Columns.Count).End(xlToLeft).Column
    oted_row = otrd_ws.Cells(Rows.Count, 1).End(xlUp).Row
    With otrd_ws.Range(Cells(3, 1), Cells(oted_row, oted_col)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = 12566463
    End With

    Application.DisplayAlerts = False
    otrd_wb.Worksheets.Add after:=otrd_wb.Worksheets(Worksheets.Count)
    Set index_ws = otrd_wb.Worksheets(Worksheets.Count)
    index_ws.Name = "����"
    index_ws.Rows(1).RowHeight = 14.5: index_ws.Rows(2).RowHeight = 14.5
    With index_ws.Range(Cells(1, 1), Cells(2, 8))
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = 3684410
        .Font.Color = 16777215
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With index_ws.Range(Cells(1, 1), Cells(2, 2))
        .MergeCells = True
        .Value = "��ԍ�"
    End With
    With index_ws.Range(Cells(1, 3), Cells(2, 3))
        .MergeCells = True
        .Value = "���x��"
    End With
    index_ws.Columns(3).NumberFormat = "@"
    With index_ws.Range(Cells(1, 4), Cells(2, 4))
        .MergeCells = True
        .Value = "�ݖ�"
    End With
    With index_ws.Range(Cells(1, 5), Cells(2, 5))
        .MergeCells = True
        .Value = "�񓚌`��"
    End With
    With index_ws.Range(Cells(1, 6), Cells(2, 6))
        .MergeCells = True
        .Value = "�I������"
    End With
    With index_ws.Range(Cells(1, 7), Cells(2, 7))
        .MergeCells = True
        .Value = "�I������"
    End With
    index_ws.Columns(8).NumberFormat = "@"
    With index_ws.Range(Cells(1, 8), Cells(2, 8))
        .MergeCells = True
        .Value = "�I�������e"
    End With
    
    index_ws.Columns(1).ColumnWidth = 4
    index_ws.Columns(2).ColumnWidth = 4
    index_ws.Columns(3).ColumnWidth = 8
    index_ws.Columns(4).ColumnWidth = 70
    index_ws.Columns(5).ColumnWidth = 8
    index_ws.Columns(6).ColumnWidth = 8
    index_ws.Columns(7).ColumnWidth = 8
    index_ws.Columns(8).ColumnWidth = 35
    index_ws.Rows.RowHeight = 14.5
    index_cnt = 0

    Load Form_Progress
    Form_Progress.StartUpPosition = 2
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - ���[�f�[�^�t�@�C���̍쐬"
    Form_Progress.Repaint
    progress_msg = "���[�f�[�^�t�@�C���̍쐬���L�����Z�����܂����B"
    Application.Visible = False
    AppActivate Form_Progress.Caption
    
    alrt_msg = "�ݒ��ʖ��o�^��QCODE������܂��B" & vbCrLf & "��肪�������m�F���ĉ������B" & vbCrLf & vbCrLf
    For i = 1 To oted_col
        DoEvents
        Form_Progress.Label1.Caption = Int(i / oted_col * 100) & "%"
        Form_Progress.Label2.Caption = "���[�f�[�^�t�@�C���쐬��" & Status_Dot(i)
        
        'OT�f�[�^��A���菇�ɂP�s�ڂ̃��x�����ݒ��ʂɂ��邩�J�E���g
        code_cnt = WorksheetFunction.CountIf(ws_setup.Columns(1), otrd_ws.Cells(1, i))
        If code_cnt = 0 And Len(otrd_ws.Cells(1, i)) >= 1 Then
            '������Ȃ��ꍇ�͂���
            alrt_msg = alrt_msg & "���[�f�[�^" & i & "��ڂ�QCODE," & otrd_ws.Cells(1, i).Value & vbCrLf
            s_r = 1
        Else
            s_r = WorksheetFunction.Match(otrd_ws.Cells(1, i), ws_setup.Columns(1), 0)
        End If
        '�`�����Ƃɂ��ꂼ�ꏈ�����{
        '�@���[�f�[�^�쐬
        otrd_ws.Activate
        Select Case Left(ws_setup.Cells(s_r, 9).Value, 1)
            Case "C", "S", "R", "H"
                otrd_ws.Columns(i).ColumnWidth = 8
                otrd_ws.Range(Cells(1, i), Cells(2, i)).MergeCells = True
            Case "M", "L"
                ct_cnt = ws_setup.Cells(s_r, 16).Value
                ma_ed = i + ct_cnt - 1
                otrd_ws.Columns(i).ColumnWidth = 3
                If otrd_ws.Cells(1, i).Value <> "" Then
                    With otrd_ws.Range(Cells(3, i), Cells(oted_row, ma_ed)).Borders(xlInsideVertical)
                        .LineStyle = xlDot
                        .Weight = xlHairline
                        .Color = 12566463
                    End With
                    otrd_ws.Range(Cells(1, i), Cells(1, ma_ed)).MergeCells = True
                    '�o�͌`�Ԃɂ��u������
                    Select Case otpt_type
                        Case "a"
                            For j = 3 To oted_row
                                If WorksheetFunction.Sum(otrd_ws.Range(Cells(j, i), Cells(j, ma_ed))) >= 1 Then
                                    With otrd_ws.Range(Cells(j, i), Cells(j, ma_ed))
                                        .Replace What:="", Replacement:="0", lookat:=xlWhole
                                    End With
                                End If
                            Next j
                        Case "b"
                            For j = 3 To oted_row
                                If WorksheetFunction.CountIf(otrd_ws.Range(Cells(j, i), Cells(j, ma_ed)), 0) >= 1 Then
                                    With otrd_ws.Range(Cells(j, i), Cells(j, ma_ed))
                                        .Replace What:="0", Replacement:="", lookat:=xlWhole
                                    End With
                                End If
                            Next j
                        Case Else
                    End Select
                End If
            Case "F"
                ma_ed = i + ct_cnt - 1
                otrd_ws.Columns(i).ColumnWidth = 15
                otrd_ws.Cells(2, i).ShrinkToFit = True
                If ma_stamp <> Format(otrd_ws.Cells(1, i).Value, "") Then
                    otrd_ws.Range(Cells(1, i), Cells(1, ma_ed)).MergeCells = True
                    ma_stamp = Format(otrd_ws.Cells(1, i).Value, "")
                End If
            Case "O"
                otrd_ws.Columns(i).ColumnWidth = 30
                otrd_ws.Range(Cells(1, i), Cells(2, i)).MergeCells = True
            Case Else
                otrd_ws.Columns(i).ColumnWidth = 8
                otrd_ws.Range(Cells(1, i), Cells(2, i)).MergeCells = True
        End Select
        '�A�����쐬
        index_ws.Activate
        index_cnt = index_cnt + 1
        lbl_adr = otrd_ws.Cells(1, i).Address(True, True)
        adr_arr = Split(lbl_adr, "$")
        lbl_adr = adr_arr(1)
        index_ws.Cells(2 + index_cnt, 1).Value = lbl_adr
        index_ws.Cells(2 + index_cnt, 2).Value = i
        index_ws.Cells(2 + index_cnt, 3).Value = otrd_ws.Cells(1, i).Value
        If s_r > 1 Then
            index_ws.Cells(2 + index_cnt, 4).Value = ws_setup.Cells(s_r, 18).Value
            Select Case Left(ws_setup.Cells(s_r, 9).Value, 1)
                Case "C"
                    index_ws.Cells(2 + index_cnt, 5).Value = "Code"
                Case "S"
                    index_ws.Cells(2 + index_cnt, 5).Value = "SA"
                    index_ws.Cells(2 + index_cnt, 6).Value = ws_setup.Cells(s_r, 16).Value
                    If ws_setup.Cells(s_r, 16).Value >= 1 Then
                        index_ws.Cells(2 + index_cnt, 6).Select
                        With Selection
                            .Value = ws_setup.Cells(s_r, 16).Value
                            For j = 1 To ws_setup.Cells(s_r, 16).Value
                                .Offset(j - 1, 1).Value = j - ws_setup.Cells(s_r, 17).Value
                                .Offset(j - 1, 2).Value = ws_setup.Cells(s_r, 19 + j - 1).Value
                            Next j
                        End With
                    ElseIf ws_setup.Cells(s_r, 19).Value <> "" Then
                        index_ws.Cells(2 + index_cnt, 8).Value = ws_setup.Cells(s_r, 19).Value
                    End If
                    index_cnt = index_cnt + ws_setup.Cells(s_r, 16).Value - 1
                Case "M", "L"
                    index_ws.Cells(2 + index_cnt, 5).Value = "MA"
                    index_ws.Cells(2 + index_cnt, 6).Value = ws_setup.Cells(s_r, 16).Value
                    If ws_setup.Cells(s_r, 16).Value >= 1 Then
                        index_ws.Cells(2 + index_cnt, 6).Select
                        With Selection
                            .Value = ws_setup.Cells(s_r, 16).Value
                            For j = 1 To ws_setup.Cells(s_r, 16).Value
                                .Offset(j - 1, 1).Value = j - ws_setup.Cells(s_r, 17).Value
                                .Offset(j - 1, 2).Value = ws_setup.Cells(s_r, 19 + j - 1).Value
                            Next j
                        End With
                    ElseIf ws_setup.Cells(s_r, 19).Value <> "" Then
                        index_ws.Cells(2 + index_cnt, 8).Value = ws_setup.Cells(s_r, 19).Value
                    End If
                    index_cnt = index_cnt + ws_setup.Cells(s_r, 16).Value - 1
                Case "R", "H"
                    index_ws.Cells(2 + index_cnt, 5).Value = "RA"
                Case "F"
                    index_ws.Cells(2 + index_cnt, 5).Value = "FA"
                    If ws_setup.Cells(s_r, 16).Value >= 1 Then
                        index_ws.Cells(2 + index_cnt, 6).Select
                        With Selection
                            .Value = ws_setup.Cells(s_r, 16).Value
                            For j = 1 To ws_setup.Cells(s_r, 16).Value
                                .Offset(j - 1, 1).Value = j - ws_setup.Cells(s_r, 17).Value
                                .Offset(j - 1, 2).Value = ws_setup.Cells(s_r, 19 + j - 1).Value
                            Next j
                        End With
                    ElseIf ws_setup.Cells(s_r, 19).Value <> "" Then
                        index_ws.Cells(2 + index_cnt, 8).Value = ws_setup.Cells(s_r, 19).Value
                    End If
                    index_cnt = index_cnt + ws_setup.Cells(s_r, 16).Value - 1
                Case "O"
                    index_ws.Cells(2 + index_cnt, 5).Value = "FA"
                Case Else
            End Select
        End If
    Next i
    Application.Visible = True
    Unload Form_Progress
    
    Application.ScreenUpdating = True
    index_ws.Activate
    ActiveWindow.FreezePanes = False
    index_ws.Cells(3, 1).Select
    ActiveWindow.FreezePanes = True
    index_cnt = index_cnt + 2
    With index_ws.Range(Cells(3, 1), Cells(index_cnt, 8))
        .Borders.LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = 14277081
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With
    With index_ws.Range(Cells(3, 1), Cells(index_cnt, 2))
        .Borders(xlInsideVertical).Color = 8421504
        .Borders(xlInsideVertical).LineStyle = xlDot
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ShrinkToFit = True
    End With
    index_ws.Columns(4).ColumnWidth = index_ws.Columns(4).ColumnWidth - 10
    index_ws.Range(Cells(3, 4), Cells(index_cnt, 4)).WrapText = True
    index_ws.Range(Cells(3, 8), Cells(index_cnt, 8)).WrapText = True
    With index_ws.Range(Cells(3, 5), Cells(index_cnt, 6))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Cells.Select
    With Selection
        .Font.Name = "Takao�S�V�b�N"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    For i = 3 To index_cnt
        index_ws.Rows(i).RowHeight = index_ws.Rows(i).RowHeight + 10
    Next i
    index_ws.Columns(4).ColumnWidth = index_ws.Columns(4).ColumnWidth + 10
    index_ws.ResetAllPageBreaks
    index_ws.VPageBreaks.Add index_ws.Cells(1, 9)
    With index_ws.PageSetup
        .RightHeader = "&P"
        .PrintTitleRows = "$1:$2"
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0)
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
        .Zoom = False
        .FitToPagesTall = False
        .FitToPagesWide = 1
    End With
    index_ws.Cells(1, 1).Select
    otrd_ws.Activate
    With otrd_ws.Range(Cells(1, 1), Cells(2, oted_col))
        .Font.Name = "Takao�S�V�b�N"
        .Font.Size = 11
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = 12566463
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    '�I�����\�L�̒���===============================
    Select Case ctz_type
        Case 1 '�I�����ԍ��ŏo��
            '�����s�v
        Case 2 '�I�������e�ŏo��
            For i = 1 To oted_col
                Select Case Left(otrd_ws.Cells(4, i).Value, 1)
                    Case "S"
                        Set fc_rng = index_ws.Columns(3).Find(What:=otrd_ws.Cells(1, i).Value)
                        If Not fc_rng Is Nothing Then
                            fc_rw = fc_rng.Row
                            For j = fc_rw To fc_rw + index_ws.Cells(fc_rw, 6).Value - 1
                                ct_no = index_ws.Cells(j, 7).Value
                                ct_label = index_ws.Cells(j, 8).Value
                                otrd_ws.Range(Cells(7, i), Cells(oted_row, i)).Replace _
                                    What:=ct_no, Replacement:=ct_label, lookat:=xlWhole
                            Next j
                        End If
                    
                    Case "M", "L"
                        Set fc_rng = index_ws.Columns(3).Find(What:=otrd_ws.Cells(1, i).Value)
                        If Not fc_rng Is Nothing Then
                            fc_rw = fc_rng.Row
                            For j = fc_rw To fc_rw + index_ws.Cells(fc_rw, 6).Value - 1
                                ct_label = index_ws.Cells(j, 8).Value
                                otrd_ws.Range(Cells(7, i + j - fc_rw), Cells(oted_row, i + j - fc_rw)).Replace _
                                    What:=1, Replacement:=ct_label, lookat:=xlWhole
                            Next j
                        End If
                
                    Case Else
                End Select
            Next i
            
        Case 3 '�ԍ��{���e�ŏo��
            For i = 1 To oted_col
                Select Case Left(otrd_ws.Cells(4, i).Value, 1)
                    Case "S"
                        Set fc_rng = index_ws.Columns(3).Find(What:=otrd_ws.Cells(1, i).Value)
                        If Not fc_rng Is Nothing Then
                            fc_rw = fc_rng.Row
                            ct_len = Len(index_ws.Cells(fc_rw, 6).Value)
                            For j = 1 To ct_len
                                zero = zero & "0"
                            Next j
                            For j = fc_rw To fc_rw + index_ws.Cells(fc_rw, 6).Value - 1
                                ct_no = index_ws.Cells(j, 7).Value
                                ct_label = Format(ct_no, zero) & "�D" & index_ws.Cells(j, 8).Value
                                otrd_ws.Range(Cells(7, i), Cells(oted_row, i)).Replace _
                                    What:=ct_no, Replacement:=ct_label, lookat:=xlWhole
                            Next j
                            zero = ""
                        End If
                    
                    Case "M", "L"
                        Set fc_rng = index_ws.Columns(3).Find(What:=otrd_ws.Cells(1, i).Value)
                        If Not fc_rng Is Nothing Then
                            fc_rw = fc_rng.Row
                            ct_len = Len(index_ws.Cells(fc_rw, 6).Value)
                            For j = 1 To ct_len
                                zero = zero & "0"
                            Next j
                            For j = fc_rw To fc_rw + index_ws.Cells(fc_rw, 6).Value - 1
                                ct_no = index_ws.Cells(j, 7).Value
                                ct_label = Format(ct_no, zero) & "�D" & index_ws.Cells(j, 8).Value
                                otrd_ws.Range(Cells(7, i + j - fc_rw), Cells(oted_row, i + j - fc_rw)).Replace _
                                    What:=1, Replacement:=ct_label, lookat:=xlWhole
                            Next j
                            zero = ""
                        End If
                
                    Case Else
                End Select
            Next i
            
        Case Else
    End Select

    '===============================================
    If otrd_ws.Cells(5, 1).Value = "Low" And otrd_ws.Cells(6, 1).Value = "High" Then
        otrd_ws.Range(Rows(3), Rows(6)).Delete
        ActiveWindow.FreezePanes = False
        otrd_ws.Cells(3, 2).Select
        ActiveWindow.FreezePanes = True
    Else
        ActiveWindow.FreezePanes = False
        otrd_ws.Cells(7, 2).Select
        ActiveWindow.FreezePanes = True
    End If
    
    '�w�b�_�̒���===================================
    If hdr_type = 1 Then  '�w�b�_1�s
        For i = 1 To oted_col
            If otrd_ws.Cells(2, i).MergeArea.Row = 2 Then
                otrd_ws.Cells(1, i).MergeCells = False
                otrd_ws.Columns(i).ColumnWidth = 8
                If otrd_ws.Cells(1, i).Value <> "" Then
                    ct_label = otrd_ws.Cells(1, i).Value
                    otrd_ws.Cells(1, i).Value = otrd_ws.Cells(1, i).Value & "_CT" & _
                        Replace(Str(otrd_ws.Cells(2, i).Value), " ", "")
                Else
                    With otrd_ws.Cells(1, i)
                        .Value = ct_label & "_CT" & Replace(Str(otrd_ws.Cells(2, i).Value), " ", "")
                        .Interior.Color = otrd_ws.Cells(1, i - 1).Interior.Color
                        .Font.Color = otrd_ws.Cells(1, i - 1).Font.Color
                    End With
                End If
            End If
        Next i
        otrd_ws.Rows(2).Delete
        otrd_ws.Rows(1).RowHeight = 29
    '�w�b�_2�s�̏ꍇ�͏�������
    End If
    '===============================================
    
    otrd_ws.Cells(1, 1).Select
    otrd_wb.SaveAs Filename:=file_path & "\1_DATA\" & rd_fn
    otrd_wb.Close
    
    Set otrd_ws = Nothing
    Set index_ws = Nothing
    Set otrd_wb = Nothing
    
    If alrt_msg <> "�ݒ��ʖ��o�^��QCODE������܂��B" & vbCrLf & "��肪�������m�F���ĉ������B" & vbCrLf & vbCrLf Then
        MsgBox alrt_msg
    End If
    
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
      ws_mainmenu.Cells(41, 6) = "11"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 11"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - ���[�f�[�^�t�@�C���̍쐬�F�Ώۃt�@�C���m" & ot_fn & "�n"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "���[�f�[�^���o�͂��܂����B", vbInformation, "MCS 2020 - RD_Creation"
End Sub

