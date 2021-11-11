Attribute VB_Name = "Module01"
Option Explicit

Sub Initial_Setting()
    Dim fd As String
    Dim log_file As String
    Dim i_cnt As Long
'--------------------------------------------------------------------------------------------------'
'�@�����ݒ菈���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2017.04.10�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Application.StatusBar = "�����ݒ� ������..."
    Application.ScreenUpdating = False
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "���C�����j���[�̋Ɩ��R�[�h�������͂ł��B", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "���C�����j���[�̍�ƃh���C�u�������͂ł��B", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    ChDrive "H"

' �e�T�u�t�H���_�̍쐬
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col)
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    End If
        
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI"
    End If
        
' 2020/4/3 - �ǋL�Fcov�t�@�C���i����p�W�v�\�t�@�C���̕\���e���v���t�@�C���j�̃R�s�[
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_cov.xlsx") = "" Then
      FileCopy "C:\MCS2020\cov.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_cov.xlsx"
    End If

    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
        Kill ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini"
    End If
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Output As #1
    Print #1, ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    Print #1, "J-FONT=���S�V�b�N"
    Print #1, "J-FONT-SIZE=8"
    Print #1, "E-FONT=Arial"
    Print #1, "E-FONT-SIZE=9"
    Print #1, "TOTAL-COLOR=204,255,255"
    Print #1, "BORDER-COLOR=128,128,128"
    Print #1, ws_mainmenu.Cells(3, 32)
    Print #1, ws_mainmenu.Cells(4, 32)
    Print #1, ws_mainmenu.Cells(5, 32)
    Close #1
'
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\1_DATA"
    End If
'
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\2_P-DATA", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\2_P-DATA"
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\2_P-DATA\YYYYMMDD PC"
    End If
'
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD"
    End If
'
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG"
    Else
        If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG\*.*") <> "" Then
            Kill ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG\*.*"
        End If
    End If
'
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\6_�[�i��", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\6_�[�i��"
    End If

' 2020/5/19 - �ǋL�F�e��ݒ�t�@�C���̃R�s�[
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_���H�w��.xlsm") = "" Then
      FileCopy "C:\MCS2020\_���H�w��.xlsm", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_���H�w��.xlsm"
    End If
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�C���w��.xlsx") = "" Then
      FileCopy "C:\MCS2020\_�C���w��.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�C���w��.xlsx"
    End If
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx") = "" Then
      FileCopy "C:\MCS2020\_�ݒ���.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx"
    End If
'
    ' �ݒ��ʂ��N���A����O��MCS�{�̂̐ݒ��ʂ�CSV�`���ŕۑ�
    Application.DisplayAlerts = False
    wb.Activate
    ws_setup.Select
    Range("A1:A2").Select
    If ws_setup.Cells(3, 1) <> "" Then
        If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG\setup", vbDirectory) = "" Then
            MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG\setup"
        End If
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\4_LOG\setup\" & Format(Now, "yyyymmddhhmmss") & "_mcs.csv", FileFormat:=xlCSV, CreateBackup:=False
        ActiveWindow.Close
    End If
    Application.DisplayAlerts = True

' �ݒ��ʂ̃N���A
    wb.Activate
    ws_setup.Select
    Rows("4:4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.ClearContents
    Range("I3").Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ws_setup.Cells(3, 1).Select
    ws_mainmenu.Select
    
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(initial_row, initial_col) = "// �����ݒ�ς݁F" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""

' �V�X�e�����O�̏o��
    ' 2020.6.3 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(41, 6) = "�����ݒ�"
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Append As #1
    Close #1
    If Err.Number > 0 Then
        Close #1
    End If
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & "*.his" <> "" Then
        Kill ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\4_LOG\" & "*.his"
    End If
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\4_LOG\" & ws_mainmenu.Cells(gcode_row, gcode_col) & ".his" For Output As #1
    Print #1, ws_mainmenu.Cells(gcode_row, gcode_col) & " MCS 2020 operation history"
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - �����ݒ芮��"
    Close #1
    MsgBox "�����ݒ肪�������܂����B", vbInformation, "MCS 2020 - Initial_Setting"
    Shell "C:\Windows\Explorer.exe " & ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbNormalFocus
    Call Finishing_Mcs2017
End Sub

Sub Setup_save()
    Dim save_rc As Integer
'--------------------------------------------------------------------------------------------------'
'�@�ݒ��ʃZ�[�u�����@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.06.26�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.03�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "���C�����j���[�̋Ɩ��R�[�h�������͂ł��B", vbExclamation, "MCS 2020 - Setup_save"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "���C�����j���[�̍�ƃh���C�u�������͂ł��B", vbExclamation, "MCS 2020 - Setup_save"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    Call Setup_Hold
    Call Filepath_Get
    
    Application.DisplayAlerts = False
    wb.Activate
    If Dir(file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx") <> "" Then
        
        Open file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx" For Append As #1
        Close #1
        If Err.Number > 0 Then
            Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx").Close
        End If
    
        ' �ۑ��̑O��FD�t�H���_���̐ݒ��ʂ�CSV�`���ŕۑ�
        Workbooks.Open Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx"
        If ActiveSheet.Cells(3, 1) <> "" Then
            If Dir(file_path & "\4_LOG\setup", vbDirectory) = "" Then
                MkDir file_path & "\4_LOG"
                MkDir file_path & "\4_LOG\setup"
            End If
            ActiveSheet.Copy
            ActiveWorkbook.SaveAs Filename:=file_path & "\4_LOG\setup\" & Format(Now, "yyyymmddhhmmss") & "_FD.csv", FileFormat:=xlCSV, CreateBackup:=False
            ActiveWindow.Close
        End If
        Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx").Close
        
        ws_setup.Select
        save_rc = MsgBox(file_path & "\3_FD�t�H���_���ɂ���" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx���㏑�����܂����B", vbYesNo + vbQuestion, "MCS 2020 - Setup_save")
        If save_rc = vbYes Then
            ActiveSheet.Copy
            ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx"
            ActiveWindow.Close
        Else
            ws_mainmenu.Select
            End
        End If
    Else
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx"
        ActiveWindow.Close
    End If
    Application.DisplayAlerts = True
    
    wb.Activate
    ws_mainmenu.Select
    
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(initial_row, initial_col) = "// �ۑ����������F" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    ' 2020.6.3 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "Save"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > Save"
    End If
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    Call Setup_Check
    Call Finishing_Mcs2017
    MsgBox "�ݒ��ʂ̓��e��ۑ����܂����B", vbInformation, "MCS 2020 - Setup_save"
End Sub

Sub Setup_load()
    Dim mcs_ini(10) As String
    Dim ini_cnt As Integer
'--------------------------------------------------------------------------------------------------'
'�@�ݒ��ʃ��[�h�����@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.06.26�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2019.07.30�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "���C�����j���[�̋Ɩ��R�[�h�������͂ł��B", vbExclamation, "MCS 2020 - Setup_load"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "���C�����j���[�̍�ƃh���C�u�������͂ł��B", vbExclamation, "MCS 2020 - Setup_load"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    Call Setup_Hold
    Call Filepath_Get
    
    Application.DisplayAlerts = False
    
    ' �ǂݍ��݂̑O��MCS�{�̂̐ݒ��ʂ�CSV�`���ŕۑ�
    wb.Activate
    ws_setup.Select
    Range("A1:A2").Select
    If ws_setup.Cells(3, 1) <> "" Then
        If Dir(file_path & "\4_LOG\setup", vbDirectory) = "" Then
            MkDir file_path & "\4_LOG"
            MkDir file_path & "\4_LOG\setup"
        End If
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=file_path & "\4_LOG\setup\" & Format(Now, "yyyymmddhhmmss") & "_mcs.csv", FileFormat:=xlCSV, CreateBackup:=False
        ActiveWindow.Close
    End If
    
    ' �ǂݍ��݃t�@�C�����I�[�v��
    Workbooks.Open Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx"
    Cells.Select
    Selection.Copy
    
    ' �ǂݍ��݃t�@�C����\��t��
    wb.Activate
    ws_setup.Select
    Range("A1:A2").Select
    ActiveSheet.Paste
    Range("A1:A2").Select
    Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_�ݒ���.xlsx").Close
    
    Application.DisplayAlerts = True
    
    wb.Activate
    ws_mainmenu.Select
      
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
     "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
        Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & _
         "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Input As #1
        ini_cnt = 1
        Do Until EOF(1)
            DoEvents
            Line Input #1, mcs_ini(ini_cnt)
            Select Case ini_cnt
            Case 8
                ws_mainmenu.Cells(3, 32) = mcs_ini(ini_cnt)
            Case 9
                ws_mainmenu.Cells(4, 32) = mcs_ini(ini_cnt)
            Case 10
                ws_mainmenu.Cells(5, 32) = mcs_ini(ini_cnt)
            End Select
            ini_cnt = ini_cnt + 1
        Loop
        Close #1
    End If
      
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(initial_row, initial_col) = "// �ǂݍ��񂾓����F" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    ' 2020.6.3 - �ǉ�
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "Load"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > Load"
    End If
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    Call Finishing_Mcs2017
    MsgBox "�ݒ��ʂ̓��e��ǂݍ��݂܂����B", vbInformation, "MCS 2020 - Setup_load"
End Sub

Sub across_wiki()
'--------------------------------------------------------------------------------------------------'
'�@�u���E�U�̋N���@�`�����ē`���ց`�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2018.07.05�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2018.07.xx�@'
'--------------------------------------------------------------------------------------------------'
    Dim objWSH As Object
    Const URL = "https://www.across-net.co.jp/across-wiki/"

    Set objWSH = CreateObject("WScript.Shell")
    objWSH.Run URL, 1
    Set objWSH = Nothing
End Sub

Sub workfolder_open()
'--------------------------------------------------------------------------------------------------'
'�@��ƃt�H���_�̕\���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'
'--------------------------------------------------------------------------------------------------'
'�@�쐬�ҁ@�@�@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�쐬���@2020.06.05�@'
'�@�ŏI�ҏW�ҁ@�e��@�m�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�ҏW���@2020.06.xx�@'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Application.ScreenUpdating = False
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "���C�����j���[�̋Ɩ��R�[�h�������͂ł��B", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "���C�����j���[�̍�ƃh���C�u�������͂ł��B", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    Application.ScreenUpdating = True
    
    Shell "C:\Windows\Explorer.exe " & ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbNormalFocus
End Sub

