Attribute VB_Name = "Module01"
Option Explicit

Sub Initial_Setting()
    Dim fd As String
    Dim log_file As String
    Dim i_cnt As Long
'--------------------------------------------------------------------------------------------------'
'　初期設定処理　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2017.04.10　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Application.StatusBar = "初期設定 処理中..."
    Application.ScreenUpdating = False
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "メインメニューの業務コードが未入力です。", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "メインメニューの作業ドライブが未入力です。", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    ChDrive "H"

' 各サブフォルダの作成
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col)
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    End If
        
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI"
    End If
        
' 2020/4/3 - 追記：covファイル（印刷用集計表ファイルの表紙テンプレファイル）のコピー
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_cov.xlsx") = "" Then
      FileCopy "C:\MCS2020\cov.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_cov.xlsx"
    End If

    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini") <> "" Then
        Kill ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini"
    End If
    Open ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\5_INI\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_mcs.ini" For Output As #1
    Print #1, ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS"
    Print #1, "J-FONT=游ゴシック"
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
    fd = Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\6_納品物", vbDirectory)
    If fd = "" Then
        MkDir ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\6_納品物"
    End If

' 2020/5/19 - 追記：各種設定ファイルのコピー
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_加工指示.xlsm") = "" Then
      FileCopy "C:\MCS2020\_加工指示.xlsm", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_加工指示.xlsm"
    End If
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_修正指示.xlsx") = "" Then
      FileCopy "C:\MCS2020\_修正指示.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_修正指示.xlsx"
    End If
    
    If Dir(ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx") = "" Then
      FileCopy "C:\MCS2020\_設定画面.xlsx", ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx"
    End If
'
    ' 設定画面をクリアする前にMCS本体の設定画面をCSV形式で保存
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

' 設定画面のクリア
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
    ws_mainmenu.Cells(initial_row, initial_col) = "// 初期設定済み：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""

' システムログの出力
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(41, 6) = "初期設定"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 初期設定完了"
    Close #1
    MsgBox "初期設定が完了しました。", vbInformation, "MCS 2020 - Initial_Setting"
    Shell "C:\Windows\Explorer.exe " & ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbNormalFocus
    Call Finishing_Mcs2017
End Sub

Sub Setup_save()
    Dim save_rc As Integer
'--------------------------------------------------------------------------------------------------'
'　設定画面セーブ処理　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.06.26　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "メインメニューの業務コードが未入力です。", vbExclamation, "MCS 2020 - Setup_save"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "メインメニューの作業ドライブが未入力です。", vbExclamation, "MCS 2020 - Setup_save"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    Call Setup_Hold
    Call Filepath_Get
    
    Application.DisplayAlerts = False
    wb.Activate
    If Dir(file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx") <> "" Then
        
        Open file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx" For Append As #1
        Close #1
        If Err.Number > 0 Then
            Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx").Close
        End If
    
        ' 保存の前にFDフォルダ内の設定画面をCSV形式で保存
        Workbooks.Open Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx"
        If ActiveSheet.Cells(3, 1) <> "" Then
            If Dir(file_path & "\4_LOG\setup", vbDirectory) = "" Then
                MkDir file_path & "\4_LOG"
                MkDir file_path & "\4_LOG\setup"
            End If
            ActiveSheet.Copy
            ActiveWorkbook.SaveAs Filename:=file_path & "\4_LOG\setup\" & Format(Now, "yyyymmddhhmmss") & "_FD.csv", FileFormat:=xlCSV, CreateBackup:=False
            ActiveWindow.Close
        End If
        Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx").Close
        
        ws_setup.Select
        save_rc = MsgBox(file_path & "\3_FDフォルダ内にある" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsxを上書きしますか。", vbYesNo + vbQuestion, "MCS 2020 - Setup_save")
        If save_rc = vbYes Then
            ActiveSheet.Copy
            ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx"
            ActiveWindow.Close
        Else
            ws_mainmenu.Select
            End
        End If
    Else
        ActiveSheet.Copy
        ActiveWorkbook.SaveAs Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx"
        ActiveWindow.Close
    End If
    Application.DisplayAlerts = True
    
    wb.Activate
    ws_mainmenu.Select
    
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    ws_mainmenu.Cells(initial_row, initial_col) = "// 保存した日時：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    ' 2020.6.3 - 追加
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
    MsgBox "設定画面の内容を保存しました。", vbInformation, "MCS 2020 - Setup_save"
End Sub

Sub Setup_load()
    Dim mcs_ini(10) As String
    Dim ini_cnt As Integer
'--------------------------------------------------------------------------------------------------'
'　設定画面ロード処理　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.06.26　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2019.07.30　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "メインメニューの業務コードが未入力です。", vbExclamation, "MCS 2020 - Setup_load"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "メインメニューの作業ドライブが未入力です。", vbExclamation, "MCS 2020 - Setup_load"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    Call Setup_Hold
    Call Filepath_Get
    
    Application.DisplayAlerts = False
    
    ' 読み込みの前にMCS本体の設定画面をCSV形式で保存
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
    
    ' 読み込みファイルをオープン
    Workbooks.Open Filename:=file_path & "\3_FD\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx"
    Cells.Select
    Selection.Copy
    
    ' 読み込みファイルを貼り付け
    wb.Activate
    ws_setup.Select
    Range("A1:A2").Select
    ActiveSheet.Paste
    Range("A1:A2").Select
    Workbooks(ws_mainmenu.Cells(gcode_row, gcode_col) & "_設定画面.xlsx").Close
    
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
    ws_mainmenu.Cells(initial_row, initial_col) = "// 読み込んだ日時：" & Format(Now, "yyyy/mm/dd hh:mm:ss")
    ws_mainmenu.Cells(initial_row, initial_col).Locked = True
    ActiveSheet.Protect Password:=""
    
    ' 2020.6.3 - 追加
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
    MsgBox "設定画面の内容を読み込みました。", vbInformation, "MCS 2020 - Setup_load"
End Sub

Sub across_wiki()
'--------------------------------------------------------------------------------------------------'
'　ブラウザの起動　～そして伝説へ～　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.07.05　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2018.07.xx　'
'--------------------------------------------------------------------------------------------------'
    Dim objWSH As Object
    Const URL = "https://www.across-net.co.jp/across-wiki/"

    Set objWSH = CreateObject("WScript.Shell")
    objWSH.Run URL, 1
    Set objWSH = Nothing
End Sub

Sub workfolder_open()
'--------------------------------------------------------------------------------------------------'
'　作業フォルダの表示　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2020.06.05　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.xx　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    Call Starting_Mcs2017
    Application.ScreenUpdating = False
    
    If ws_mainmenu.Cells(gcode_row, gcode_col) = "" Then
        MsgBox "メインメニューの業務コードが未入力です。", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gcode_row, gcode_col).Select
        Call Finishing_Mcs2017
        End
    End If
    
    If ws_mainmenu.Cells(gdrive_row, gdrive_col) = "" Then
        MsgBox "メインメニューの作業ドライブが未入力です。", vbExclamation, "MCS 2020 - Initial_Setting"
        ws_mainmenu.Cells(gdrive_row, gdrive_col).Select
        Call Finishing_Mcs2017
        End
    End If
    Application.ScreenUpdating = True
    
    Shell "C:\Windows\Explorer.exe " & ws_mainmenu.Cells(gdrive_row, gdrive_col) & ":\" & ws_mainmenu.Cells(gcode_row, gcode_col) & "\MCS", vbNormalFocus
End Sub

