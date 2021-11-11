Attribute VB_Name = "Module54"
Option Explicit
    Dim spread_fn As String
    Dim spread_fd As String
    
    Dim wb_spread As Workbook
    Dim ws_spread0 As Worksheet
    Dim ws_spread1 As Worksheet
    Dim ws_spread2 As Worksheet
    Dim ws_spread3 As Worksheet
    
    Dim wb_print As Workbook
    Dim tn_tab As Boolean

Public Sub Print_spreadsheet()
    Dim waitTime As Variant
    
    Dim rc As Integer
    Dim yen_pos As Long
    
    Dim m_area As Range, tab_no_fc As Range
    Dim head_split As Boolean

    Dim st_rw As Long, ed_rw As Long, usd_rw As Long, gt_rw As Long, _
    dmy_rw As Long, del_rw As Long, top_rw As Long, del_cl As Long, _
    st_cl As Long, ed_cl As Long, height_lim As Long, height_sum As Long, _
    wrap_cnt As Long, ctz_cnt As Long, cpy_cnt As Long, top_ctz As Long, _
    cpy As Long, wrap As Long, del_st As Long, del_ed As Long, s As Long, t As Long

    Dim rng_add As String, face_label As String, print_label As String, _
    table_label As String, not_found As String

    Dim max_cnt As Long, hyo_cnt As Long

    Dim p_cnt As Long, p_temp As Long
    Dim max_row As Long
'--------------------------------------------------------------------------------------------------'
'　印刷用集計表ファイルの作成  　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　'
'--------------------------------------------------------------------------------------------------'
'　作成者　　　田中義晃　　　　　　　　　　　　　　　　　　　　　　　　　　　　作成日　2018.04.04　'
'　最終編集者　菊崎　洋　　　　　　　　　　　　　　　　　　　　　　　　　　　　編集日　2020.06.03　'
'--------------------------------------------------------------------------------------------------'
    On Error Resume Next
    
    Call Starting_Mcs2017
    Call Filepath_Get
    Call Setup_Check

    wb.Activate
    ws_mainmenu.Select
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\5_INI"
    
    If Dir(file_path & "\5_INI\" & ws_mainmenu.Cells(3, 8) & "cov.xlsx") = "" Then
        MsgBox "表紙テンプレートファイル［*cov.xlsx］がみつかりません。" _
         & vbCrLf & "5_INIフォルダ内に表紙テンプレートファイルを用意してください。", vbExclamation, "MCS 2020 - Cover_Procedure"
        Call Finishing_Mcs2017
        End
    End If
    
    ChDrive file_path
    ChDir file_path
    ChDir file_path & "\SUM"
    
    rc = MsgBox("印刷設定をする集計表Excelファイルが必要です。" _
     & vbCrLf & "集計表Excelファイルはありますか？" & vbCrLf & vbCrLf _
     & "「はい」　→ すでにある集計表Excelファイルを選択" & vbCrLf & "「いいえ」→ 集計表Excelファイルを作成する", _
     vbYesNoCancel + vbQuestion, "集計表Excelファイル作成の確認")
    If rc = vbNo Then
'ウザイんで表示しないようにした - 2018/6/21
'        MsgBox "集計表Excelファイルを作成します。集計サマリーデータを選択してください。", , "MCS 2020 - Print_spreadsheet"
        Call Spreadsheet_Creation
    ElseIf rc = vbCancel Then
        wb.Activate
        ws_mainmenu.Select
        End
    End If

'ウザイんで表示しないようにした - 2018/6/21
'    MsgBox "印刷設定をする集計表Excelファイルを選択してください。", , "MCS 2020 - Print_spreadsheet"

step00:
    spread_fn = Application.GetOpenFilename("集計表Excelファイル,*.xlsx", , "集計表Excelファイルを開く")
    If spread_fn = "False" Then
        ' キャンセルボタンの処理
        wb.Activate
        ws_mainmenu.Select
        End
    ElseIf spread_fn = "" Then
        MsgBox "集計表Excelファイルを選択してください。", vbExclamation, "MCS 2020 - Print_spreadsheet"
        wb.Activate
        ws_mainmenu.Select
        GoTo step00
    ElseIf InStr(spread_fn, "_集計表") = 0 Then
        MsgBox "集計表Excelファイルを選択してください。", vbExclamation, "MCS 2020 - Print_spreadsheet"
        wb.Activate
        ws_mainmenu.Select
        GoTo step00
    End If

    Workbooks.Open spread_fn
    ' フルパスからフォルダ名の取得
    yen_pos = InStrRev(spread_fn, "\")
    spread_fd = Left(spread_fn, yen_pos - 1)
    
    spread_fd = spread_fd & "\印刷用"
    If Dir(spread_fd, vbDirectory) = "" Then
        MkDir spread_fd
    End If
    
    ' フルパスからファイル名の取得
    spread_fn = Dir(spread_fn)
    
    Set wb_spread = Workbooks(spread_fn)
    Set ws_spread0 = wb_spread.Worksheets(1)
    Set ws_spread1 = wb_spread.Worksheets(2)
    Set ws_spread2 = wb_spread.Worksheets(3)
    Set ws_spread3 = wb_spread.Worksheets(4)

    max_cnt = WorksheetFunction.CountA(ws_spread0.Columns(1)) - 1

'2018/06/19 - 追記 ==========================
    Application.ScreenUpdating = False
    Load Form_Progress
    Form_Progress.StartUpPosition = 1
    Form_Progress.Show vbModeless
    Form_Progress.Caption = "MCS 2020 - 印刷用集計表ファイルの作成"
    Form_Progress.Repaint
    progress_msg = "印刷用集計表ファイルの作成をキャンセルしました。"
    Application.Visible = False
    AppActivate Form_Progress.Caption

    Form_Progress.Label1.Caption = "初期設定中"
    Form_Progress.Label2.Caption = "しばらくお待ちください..."
    Form_Progress.Label3.Caption = "[1/1ファイル]"
    DoEvents
'============================================

    Call paradigm_procedure

'集計表 詳細印刷設定
'----N%
    '表頭折り返し処理
    ws_spread1.Activate
    If tn_tab = True Then
' MCODE処理のときに都合が悪いのでとりあえずコメントアウト
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    If ed_rw = st_rw Then
        MsgBox "C列に表側情報が存在しない為、処理を中断します。"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP1/6 印刷用集計表ファイル（Ｎ％表）表頭折り返し処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        '表側上一番上にくるカテゴリの行を取得
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        '表頭右端の列番号を取得
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        '表頭分割分増幅し、カテゴリーをＭＥＣＥに調整
        dmy_rw = st_rw + gt_rw
        Do While ed_cl > 16
            del_rw = Cells(dmy_rw, 7).End(xlDown).Row + 1
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = Cells(dmy_rw, 7).End(xlDown).Row
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")
    
    '表側分割処理 N%表
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP2/6 印刷用集計表ファイル（Ｎ％表）表側分割処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        '表頭割れの確認
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 2
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        '表側分割有無の確認（表側項目25以上に対し実行）
        If ctz_cnt > 24 Then
            '表側分割数の計算
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            '表側分割分増幅
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 2), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 3, 1).Insert
            Next cpy
            'カテゴリーをＭＥＣＥに調整
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  '１（全体）～２３項目目の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt * 2 - 1
                            del_st = del_st + 24 * 2
                            '表側ラベル処理 ①
                            '---------------------------------------------------------------------------
                            '２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '以下は合成フェース実装後に必要な処理
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap
                    Case 2 To cpy_cnt - 1  '２４項目目からコピーした末尾の表のひとつ前の項目まで
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) * 2 - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 * 2 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt * 2 - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24 * 2, 4), Cells(del_ed + 24 * 2, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            '表側ラベル処理 ②
                            '---------------------------------------------------------------------------
                            '１項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 * 2 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24 * 2, 3).End(xlDown).Row
                        Next wrap
                    Case Is = cpy_cnt  '分割末尾の表の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) * 2 - 1
                            '表側ラベル処理 ③
                            '---------------------------------------------------------------------------
                            '２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
            
                            '表側ラベル処理 ④
                            '---------------------------------------------------------------------------
                            '１項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw * 2, 3).End(xlDown).Row
                        Next wrap
                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'----

'=====ここからN表
    ws_spread2.Activate
    If tn_tab = True Then
' MCODE処理のときに都合が悪いのでとりあえずコメントアウト
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    ' カテゴリ行の高さを12ptから24pへ変更
    st_rw = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 3).End(xlDown).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    Do
        For dmy_rw = st_rw To ed_rw
            Rows(dmy_rw).RowHeight = 24
        Next dmy_rw
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(Cells(ed_rw, 1).End(xlUp).Row, 3).End(xlDown).Row
        If ed_rw = 1 Then Exit Do
    Loop
    ' 表頭折り返し処理
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    If ed_rw = st_rw Then
        MsgBox "C列に表側情報が存在しない為、処理を中断します。"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP3/6 印刷用集計表ファイル（Ｎ表）表頭折り返し処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        '「全体」カテゴリの表示位置の垂直位置を変更
        If Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value = 0 Then
            Cells(Cells(st_rw, 3).End(xlDown).Row, 4).VerticalAlignment = xlTop
        End If
        ' 表側上一番上にくるカテゴリの行を取得
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        ' 表頭右端の列番号を取得
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        ' 表頭分割分増幅し、カテゴリーをＭＥＣＥに調整
        dmy_rw = st_rw + gt_rw
        del_rw = dmy_rw
        Do While ed_cl > 16
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    del_rw = del_rw + 1
                    Exit Do
                End If
            Loop
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = dmy_rw
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    Exit Do
                End If
            Loop
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")

    ' 表側分割処理 N表
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP4/6 印刷用集計表ファイル（Ｎ表）表側分割処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        ' 表頭割れの確認
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 1
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        ' 表側分割有無の確認（表側項目25以上に対し実行）
        If ctz_cnt > 24 Then
            ' 表側分割数の計算
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            ' 表側分割分増幅
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 1), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 2, 1).Insert
            Next cpy
            ' カテゴリーをMECEに調整
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  ' １（全体）～２３項目目の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt - 1
                            del_st = del_st + 24
                            ' 表側ラベル処理 ①
                            '---------------------------------------------------------------------------
                            ' ２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            ' 以下は合成フェース実装後に必要な処理
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap

                    Case 2 To cpy_cnt - 1  ' ２４項目目からコピーした末尾の表のひとつ前の項目まで
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24, 4), Cells(del_ed + 24, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ' 表側ラベル処理 ②
                            '---------------------------------------------------------------------------
                            ' １項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24, 3).End(xlDown).Row
                        Next wrap

                    Case Is = cpy_cnt  ' 分割末尾の表の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) - 1
                            ' 表側ラベル処理 ③
                            '---------------------------------------------------------------------------
                            ' ２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With

                            ' 表側ラベル処理 ④
                            '---------------------------------------------------------------------------
                            ' １項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw, 3).End(xlDown).Row
                        Next wrap

                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'=====ここまでN表

'=====ここから%表
    ws_spread3.Activate
    If tn_tab = True Then
' MCODE処理のときに都合が悪いのでとりあえずコメントアウト
'        Columns("D:E").Borders(xlInsideVertical).LineStyle = False
    End If
    ' カテゴリ行の高さを12ptから24pへ変更
    st_rw = Cells(Cells(Rows.Count, 1).End(xlUp).Row, 3).End(xlDown).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    Do
        For dmy_rw = st_rw To ed_rw
            Rows(dmy_rw).RowHeight = 24
        Next dmy_rw
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(Cells(ed_rw, 1).End(xlUp).Row, 3).End(xlDown).Row
        If ed_rw = 1 Then Exit Do
    Loop
    ' 表頭折り返し処理
    st_cl = 6
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    If ed_rw = st_rw Then
        MsgBox "C列に表側情報が存在しない為、処理を中断します。"
        End
    End If
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP5/6 印刷用集計表ファイル（％表）表頭折り返し処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        '「全体」カテゴリの表示位置の垂直位置を変更
        If Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value = 0 Then
            Cells(Cells(st_rw, 3).End(xlDown).Row, 4).VerticalAlignment = xlTop
        End If
        ' 表側上一番上にくるカテゴリの行を取得
        gt_rw = WorksheetFunction.Match(Cells(Cells(st_rw, 3).End(xlDown).Row, 3).Value, _
            Range(Cells(st_rw, 3), Cells(ed_rw, 3)), 0) - 1
        ' 表頭右端の列番号を取得
        ed_cl = Cells(st_rw + gt_rw, st_cl).End(xlToRight).Column
        With Cells(st_rw, 16)
            .Formula = "=if(len(A" & st_rw & ")<>0,A" & st_rw & ","""" )"
            .HorizontalAlignment = xlRight
            .Font.Color = RGB(128, 128, 128)
        End With
        ' 表頭分割分増幅し、カテゴリーをＭＥＣＥに調整
        dmy_rw = st_rw + gt_rw
        del_rw = dmy_rw
        Do While ed_cl > 16
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    del_rw = del_rw + 1
                    Exit Do
                End If
            Loop
            del_cl = Cells(dmy_rw, 7).End(xlToRight).Column
            rng_add = Replace(Str(dmy_rw - 2), " ", "") & ":" & Replace(Str(del_rw), " ", "")
            Rows(rng_add).Copy
            Cells(del_rw + 1, 1).Insert
            Range(Cells(dmy_rw - 2, 17), Cells(del_rw, del_cl)).Delete Shift:=xlToLeft
            dmy_rw = Cells(del_rw, 3).End(xlDown).Row
            del_rw = dmy_rw
            Do
                If Cells(del_rw + 1, 3).Value <> "" Then
                    del_rw = del_rw + 1
                Else
                    Exit Do
                End If
            Loop
            Range(Cells(dmy_rw - 2, 6), Cells(del_rw, 16)).Delete Shift:=xlToLeft
            ed_cl = Cells(dmy_rw, Columns.Count).End(xlToLeft).Column
        Loop
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then Exit Do
        ed_rw = ed_rw + 1
    Loop
    DoEvents
    Columns("F:P").ColumnWidth = 10.5
    Form_Progress.Label1.Caption = "100%"
    waitTime = Now + TimeValue("0:00:01")

    ' 表側分割処理 %表
    st_rw = Cells(Rows.Count, 1).End(xlUp).Row
    ed_rw = Cells(Rows.Count, 3).End(xlUp).Row
    hyo_cnt = 0
    Do
        DoEvents
        hyo_cnt = hyo_cnt + 1
        Form_Progress.Label1.Caption = Int(hyo_cnt / max_cnt * 100) & "%"
        Form_Progress.Label2.Caption = "STEP6/6 印刷用集計表ファイル（％表）表側分割処理中" & Status_Dot(ed_rw)
        Form_Progress.Label3.Caption = "[1/1ファイル]"
        
        wrap_cnt = 0
        cpy_cnt = 0
        ctz_cnt = 0
        ' 表頭割れの確認
        top_rw = Cells(st_rw, 3).End(xlDown).Row - 2
        dmy_rw = top_rw + 2
        Do
            If Cells(dmy_rw, 3) <> "" Then
                If wrap_cnt = 0 Then
                    ctz_cnt = ctz_cnt + 1
                End If
                dmy_rw = dmy_rw + 1
            Else
                wrap_cnt = wrap_cnt + 1
                dmy_rw = Cells(dmy_rw, 3).End(xlDown).Row
                If dmy_rw > ed_rw Then
                    Exit Do
                End If
            End If
        Loop
        ' 表側分割有無の確認（表側項目25以上に対し実行）
        If ctz_cnt > 24 Then
            ' 表側分割数の計算
            cpy_cnt = WorksheetFunction.RoundDown(ctz_cnt / 24, 0)
            If ctz_cnt Mod 24 > 0 Then
                cpy_cnt = cpy_cnt + 1
            End If
            ' 表側分割分増幅
            rng_add = Replace(Str(Cells(st_rw, 3).End(xlDown).Row - 2), " ", "") & ":" & Replace(Str(ed_rw + 1), " ", "")
            For cpy = 1 To cpy_cnt - 1
                Rows(rng_add).Copy
                Cells(ed_rw + 2, 1).Insert
            Next cpy
            ' カテゴリーをMECEに調整
            del_st = Cells(st_rw, 3).End(xlDown).Row
            For cpy = 1 To cpy_cnt
                Select Case cpy
                    Case 1  ' １（全体）～２３項目目の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + ctz_cnt - 1
                            del_st = del_st + 24
                            ' 表側ラベル処理 ①
                            '---------------------------------------------------------------------------
                            ' ２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_st - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            ' 以下は合成フェース実装後に必要な処理
                            If m_area.Row = 8 And Len(face_label) > 62 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 6 And Len(face_label) > 48 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 4 And Len(face_label) > 35 Then
                                m_area.Cells(1, 1).Value = ""
                            ElseIf m_area.Row = 2 And Len(face_label) > 14 Then
                                m_area.Cells(1, 1).Value = ""
                            End If
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st - 2, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With
                            del_st = Cells(del_st, 3).End(xlDown).Row
                        Next wrap

                    Case 2 To cpy_cnt - 1  ' ２４項目目からコピーした末尾の表のひとつ前の項目まで
                        For wrap = 1 To wrap_cnt
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            del_ed = del_st + 24 * (cpy - 1) - 1
                            If ctz_cnt > 24 * cpy Then
                                rng_add = Replace(Str(del_ed + 24 + 1), " ", "") & ":" & Replace(Str(del_st + ctz_cnt - 1), " ", "")
                                Rows(rng_add).Delete
                                With Range(Cells(del_ed + 24, 4), Cells(del_ed + 24, ed_cl)).Borders(xlBottom)
                                    .LineStyle = xlContinuous
                                    .Weight = xlThin
                                    .Color = RGB(128, 128, 128)
                                End With
                            End If
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ' 表側ラベル処理 ②
                            '---------------------------------------------------------------------------
                            ' １項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            Set m_area = Cells(del_st + 24 - 1, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            del_st = Cells(del_st + 24, 3).End(xlDown).Row
                        Next wrap

                    Case Is = cpy_cnt  ' 分割末尾の表の処理
                        For wrap = 1 To wrap_cnt
                            del_ed = del_st + 24 * (cpy_cnt - 1) - 1
                            ' 表側ラベル処理 ③
                            '---------------------------------------------------------------------------
                            ' ２４項目目が属する表側ラベルを取得
                            Set m_area = Cells(del_ed, 4).MergeArea
                            face_label = m_area.Cells(1, 1).Value
                            '---------------------------------------------------------------------------
                            rng_add = Replace(Str(del_st), " ", "") & ":" & Replace(Str(del_ed), " ", "")
                            Rows(rng_add).Delete
                            ed_cl = Cells(del_st, Columns.Count).End(xlToLeft).Column
                            With Range(Cells(del_st - 1, 4), Cells(del_st - 1, ed_cl)).Borders(xlBottom)
                                .LineStyle = xlContinuous
                                .Weight = xlThin
                                .Color = RGB(128, 128, 128)
                            End With

                            ' 表側ラベル処理 ④
                            '---------------------------------------------------------------------------
                            ' １項目目が属する表側ラベルを確認し、空白なら表側ラベルを上書きする
                            Set m_area = Cells(del_st, 4).MergeArea
                            If m_area.Cells(1, 1).Value = "" Then
                                If m_area.Row >= 8 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 6 And Len(face_label) <= 48 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 4 And Len(face_label) <= 35 Then
                                    m_area.Cells(1, 1).Value = face_label
                                ElseIf m_area.Row = 2 And Len(face_label) <= 14 Then
                                    m_area.Cells(1, 1).Value = face_label
                                End If
                            End If
                            '---------------------------------------------------------------------------
                            dmy_rw = ctz_cnt - (cpy - 1) * 24
                            del_st = Cells(del_st + dmy_rw, 3).End(xlDown).Row
                        Next wrap

                    Case Else
                End Select
            Next cpy
            Set m_area = Nothing
        End If
        ed_rw = Cells(st_rw, 3).End(xlUp).Row
        st_rw = Cells(st_rw, 1).End(xlUp).Row
        If st_rw = ed_rw Then
            Exit Do
        End If
    Loop
'=====ここまで%表

'---
'改頁設定
    ' N%表の改頁設定
    ws_spread1.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    ' 目次頁へのページ番号付与（先頭）
    With ws_spread0.Cells(1, 6)
        .Value = "頁"
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
    s = 1
    not_found = "Not Found" & vbCrLf
    ws_spread0.Cells(2, 6).Value = s
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            ' 目次頁へのページ番号付与
            Set tab_no_fc = ws_spread0.Columns(2).Find(What:=ws_spread1.Cells(st_rw, 1).Value, lookat:=xlWhole)
            If ws_spread0.Cells(tab_no_fc.Row, 6).Value = "" Then
                s = s + 1
            End If
            If tab_no_fc Is Nothing Then
                not_found = not_found & ws_spread1.Cells(st_rw, 1).Value & vbCrLf
            ElseIf ws_spread1.Cells(st_rw, 1).Value <> "" Then
                ws_spread0.Cells(tab_no_fc.Row, 6).Value = s
                t = ws_spread0.Cells(tab_no_fc.Row, 6).Row
                Do
                    If ws_spread0.Cells(t, 6).Offset(-1, 0).Value <> "" Then
                        Exit Do
                    Else
                        ws_spread0.Cells(t, 6).Offset(-1, 0).Value = s - 1
                    End If
                    t = ws_spread0.Cells(t, 6).Offset(-1, 0).Row
                Loop
            End If
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 4).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                            ' 改頁をページ番号に反映
                            If dmy_rw > st_rw Then
                                s = s + 1
                            End If
                            '----
                        End If
                        dmy_rw = wrap
                    End If
                    ' 最終ページの改ページ調整
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' N表の改頁設定
    ws_spread2.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 1).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                        End If
                        dmy_rw = wrap
                    End If
                    ' 最終ページの改ページ調整
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' %表の改頁設定
    ws_spread3.Activate
    ActiveWindow.View = xlPageBreakPreview
    height_lim = 760
    height_sum = 0
    st_rw = 1
    usd_rw = Cells(Rows.Count, 3).End(xlUp).Row + 1
    Do
        ed_rw = Cells(st_rw, 1).End(xlDown).Row: If ed_rw > usd_rw Then ed_rw = usd_rw + 1
        height_sum = height_sum + Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
        If height_sum > height_lim Then
            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(st_rw, 2)
            height_sum = Rows(Replace(Str(st_rw) & ":" & Str(ed_rw - 1), " ", "")).Height
            If height_sum > height_lim Then
                height_sum = 0
                For wrap = st_rw To ed_rw - 1
                    height_sum = height_sum + Rows(Replace(Str(wrap) & ":" & Str(wrap), " ", "")).Height
                    If WorksheetFunction.CountA(Range(Cells(wrap, 3), Cells(wrap, 16))) = 0 And Cells(wrap - 1, 1).Value = "" Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = Rows(Replace(Str(dmy_rw + 1) & ":" & Str(wrap), " ", "")).Height
                        End If
                        dmy_rw = wrap
                    End If
                    '最終ページの改ページ調整
                    If wrap = ed_rw - 1 Then
                        If height_sum > height_lim Then
                            ActiveWindow.SelectedSheets.HPageBreaks.Add before:=Cells(dmy_rw + 1, 2)
                            height_sum = 0
                        End If
                    End If
                Next wrap
            End If
        End If
        st_rw = ed_rw
        If st_rw > usd_rw Then
            Exit Do
        End If
    Loop
    If ws_spread0.Cells(1, 2).End(xlDown).Row > tab_no_fc.Row Then
        For t = tab_no_fc.Row To ws_spread0.Cells(1, 2).End(xlDown).Row
            ws_spread0.Cells(t, 6).Value = ws_spread0.Cells(tab_no_fc.Row, 6).Value
        Next t
    End If
    If not_found <> "Not Found" & vbCrLf Then
        MsgBox (not_found)
    End If
    Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' ページ番号の書式設定 - 2018.9.20
    ws_spread0.Activate
    ws_spread0.Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "_ * #,##0_ ;_ * -#,##0_ ;_ * ""-""_ ;_ @_ "
    ws_spread0.Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' 表№の書式設定 - 2020.3.30
    ws_spread0.Activate
    ws_spread0.Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ShrinkToFit = True
    ws_spread0.Cells(1, 1).Select
    ActiveWindow.View = xlNormalView

    ' 2020.4.2 - 出力が１ページしかないときの不具合修正
    ws_spread0.Activate
    max_row = Cells(Rows.Count, 1).End(xlUp).Row
    For p_cnt = 2 To max_row
        If Cells(p_cnt, 6) <> "" Then
            p_temp = Cells(p_cnt, 6)
        Else
            Cells(p_cnt, 6) = p_temp
        End If
    Next p_cnt

'2018/06/19 - 追記 ==========================
    Application.Visible = True
    Unload Form_Progress
'============================================

' 集計表印刷設定ここまで
'======
' 表紙設定
    Call cover_procedure

'======================================================================
' ＰＤＦ出力ここから
'======================================================================
    If tn_tab = True Then
        print_label = Replace(spread_fn, "_集計表", "_単純集計表")
    Else
        print_label = spread_fn
    End If
    wb_spread.SaveCopyAs spread_fd & "\" & "【印刷用TEMP】" & print_label
    For s = 1 To 3
        Select Case s
            Case 1
                Workbooks.Open Filename:=(spread_fd & "\" & "【印刷用TEMP】" & print_label)
                Set wb_print = Workbooks("【印刷用TEMP】" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(5).Delete
                wb_print.Worksheets(4).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\（印刷用）" & print_label)
                Application.DisplayAlerts = True
                Workbooks("（印刷用）" & print_label).Activate
                Call publish_procedure
                Workbooks("【印刷用TEMP】" & print_label).Close
                Workbooks("（印刷用）" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing

            Case 2
                Workbooks.Open Filename:=(spread_fd & "\" & "【印刷用TEMP】" & print_label)
                Set wb_print = Workbooks("【印刷用TEMP】" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(5).Delete
                wb_print.Worksheets(3).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\（件数表）" & print_label)
                Application.DisplayAlerts = True
                Workbooks("（件数表）" & print_label).Activate
                Call publish_procedure
                Workbooks("【印刷用TEMP】" & print_label).Close
                Workbooks("（件数表）" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing

            Case 3
                Workbooks.Open Filename:=(spread_fd & "\" & "【印刷用TEMP】" & print_label)
                Set wb_print = Workbooks("【印刷用TEMP】" & print_label)
                Application.DisplayAlerts = False
                wb_print.Worksheets(4).Delete
                wb_print.Worksheets(3).Delete
                Call CoverMark_procedure
                wb_print.SaveAs Filename:=(spread_fd & "\（構成比表）" & print_label)
                Application.DisplayAlerts = True
                Workbooks("（構成比表）" & print_label).Activate
                Call publish_procedure
                Workbooks("【印刷用TEMP】" & print_label).Close
                Workbooks("（構成比表）" & print_label).Close
                wb_print.Close
                Set wb_print = Nothing
            Case Else
        End Select
    Next s
    Application.DisplayAlerts = False
    wb_spread.Activate
    Kill spread_fd & "\" & "【印刷用TEMP】" & print_label
    wb_spread.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
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
    
' システムログの出力 - 2020.5.14
    ' 2020.6.3 - 追加
    ActiveSheet.Unprotect Password:=""
    ws_mainmenu.Cells(initial_row, initial_col).Locked = False
    If (Len(ws_mainmenu.Cells(41, 6)) > 70) Or (Len(ws_mainmenu.Cells(41, 6)) = 0) Then
      ws_mainmenu.Cells(41, 6) = "26"
    Else
      ws_mainmenu.Cells(41, 6) = ws_mainmenu.Cells(41, 6) & " > 26"
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 印刷用集計表ファイルの作成：対象ファイル［" & spread_fn & "］"
    Close #1
    Call Finishing_Mcs2017
    MsgBox "印刷用集計表ファイルが完成しました。", vbInformation, "MCS 2020 - Print_spreadsheet"
End Sub

Private Sub paradigm_procedure()
' 目次ページ（ws_spread0）の印刷設定
    Dim max_row As Long
    Dim i_cnt As Long
    Dim now_row As Long
    
    ws_spread0.Activate
    max_row = Cells(Rows.Count, 1).End(xlUp).Row

    ' 目次行の高さ調整
    For i_cnt = 2 To max_row
        now_row = Rows(i_cnt).RowHeight
        now_row = now_row / 12.75    ' 目次１項目の行数を算出
        now_row = (now_row / 5) * 12.75
        Rows(i_cnt).RowHeight = Rows(i_cnt).RowHeight + now_row
    Next i_cnt
    
    If WorksheetFunction.CountA(Columns("D")) - 1 = 0 Then
        tn_tab = True
    Else
        tn_tab = False
    End If
    If Cells(1, 3).Value = "MCODE" Then
        Columns(3).Delete
    End If
    With Columns("A:B")
        .ColumnWidth = 4
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Columns("C:E")
        .ColumnWidth = 42
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    With Columns("F:H")
        .ColumnWidth = 6
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    Columns(6).Insert
    With Columns(6)
        .ColumnWidth = 4
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$A:$F"
        .PrintTitleRows = "$1:$1"               ' 印刷タイトル行の指定
        .PrintHeadings = False                  ' 行列番号を含めて印刷＝しない
        .PrintGridlines = False                 ' 枠線を印刷＝しない
        .PrintComments = xlPrintNoComments      ' セルのコメントを印刷＝しない
        .PrintQuality = 600                     ' 印刷品質＝600dpi
        .CenterHorizontally = False             ' 頁中央（水平）に印刷＝しない
        .CenterVertically = False               ' 頁中央（垂直）に印刷＝しない
        .Orientation = xlLandscape              ' 印刷方向＝横
        .Draft = False                          ' 簡易印刷＝しない
        .PaperSize = xlPaperA4                  ' 頁サイズ＝A4
        .Order = xlDownThenOver                 ' 頁番号の付番規則＝上から下
        .BlackAndWhite = False                  ' 白黒印刷＝しない
        .Zoom = 100                             ' 印刷倍率
        .PrintErrors = xlPrintErrorsDisplayed   ' エラー表示の印刷＝見たまま印刷
        ' 以下､余白設定
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(0)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
    End With
    Columns("G:I").Delete
    If tn_tab = True Then
        Columns("C").ColumnWidth = 9.88
        Columns("D:E").ColumnWidth = 58
        Columns("G").ColumnWidth = 4
    End If

    
    ' 目次の最終調整
    Range("A1").Select
    Selection.CurrentRegion.Select
    With Selection.Font
        .Name = "游ゴシック"
        .Size = 8
    End With
    Rows("1:1").Select
    With Selection.Font
        .Name = "游ゴシック"
        .Size = 9
    End With
    Range("A1").Select

' 集計表ページの設定
    ' N%
    ws_spread1.Activate
    Columns("C").ColumnWidth = 3
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' 行列番号を含めて印刷＝しない
        .PrintGridlines = False                 ' 枠線を印刷＝しない
        .PrintComments = xlPrintNoComments      ' セルのコメントを印刷＝しない
        .PrintQuality = 600                     ' 印刷品質＝600dpi
        .CenterHorizontally = False             ' 頁中央（水平）に印刷＝しない
        .CenterVertically = False               ' 頁中央（垂直）に印刷＝しない
        .Orientation = xlLandscape              ' 印刷方向＝横
        .Draft = False                          ' 簡易印刷＝しない
        .PaperSize = xlPaperA4                  ' 頁サイズ＝A4
        .FirstPageNumber = 1                    ' 先頭頁番号＝1
        .Order = xlDownThenOver                 ' 頁番号の付番規則＝上から下
        .BlackAndWhite = False                  ' 白黒印刷＝しない
        .Zoom = 80                              ' 印刷倍率
        .PrintErrors = xlPrintErrorsDisplayed   ' エラー表示の印刷＝見たまま印刷
        ' 以下、余白設定
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        '以下、フッター設定
        .RightFooter = "&""Century""&9&P"
    End With
    ' N
    ws_spread2.Activate
    With Columns("C")
        .ColumnWidth = 3
        .VerticalAlignment = xlTop
    End With
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
        .VerticalAlignment = xlTop
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' 行列番号を含めて印刷＝しない
        .PrintGridlines = False                 ' 枠線を印刷＝しない
        .PrintComments = xlPrintNoComments      ' セルのコメントを印刷＝しない
        .PrintQuality = 600                     ' 印刷品質＝600dpi
        .CenterHorizontally = False             ' 頁中央（水平）に印刷＝しない
        .CenterVertically = False               ' 頁中央（垂直）に印刷＝しない
        .Orientation = xlLandscape              ' 印刷方向＝横
        .Draft = False                          ' 簡易印刷＝しない
        .PaperSize = xlPaperA4                  ' 頁サイズ＝A4
        .FirstPageNumber = 1                    ' 先頭頁番号＝1
        .Order = xlDownThenOver                 ' 頁番号の付番規則＝上から下
        .BlackAndWhite = False                  ' 白黒印刷＝しない
        .Zoom = 80                              ' 印刷倍率
        .PrintErrors = xlPrintErrorsDisplayed   ' エラー表示の印刷＝見たまま印刷
        ' 以下、余白設定
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        ' 以下、フッター設定
        .RightFooter = "&""Century""&9&P"
    End With
    ' %
    ws_spread3.Activate
    With Columns("C")
        .ColumnWidth = 3
        .VerticalAlignment = xlTop
    End With
    Columns("D").ColumnWidth = 10.63
'    If tn_tab = True Then
'        Columns("E").Insert
'    End If
    With Columns("E")
        .ColumnWidth = 35
        .ShrinkToFit = True
        .VerticalAlignment = xlTop
    End With
    With ActiveSheet.PageSetup
        .PrintArea = "$C:$P"
        .PrintHeadings = False                  ' 行列番号を含めて印刷＝しない
        .PrintGridlines = False                 ' 枠線を印刷＝しない
        .PrintComments = xlPrintNoComments      ' セルのコメントを印刷＝しない
        .PrintQuality = 600                     ' 印刷品質＝600dpi
        .CenterHorizontally = False             ' 頁中央（水平）に印刷＝しない
        .CenterVertically = False               ' 頁中央（垂直）に印刷＝しない
        .Orientation = xlLandscape              ' 印刷方向＝横
        .Draft = False                          ' 簡易印刷＝しない
        .PaperSize = xlPaperA4                  ' 頁サイズ＝A4
        .FirstPageNumber = 1                    ' 先頭頁番号＝1
        .Order = xlDownThenOver                 ' 頁番号の付番規則＝上から下
        .BlackAndWhite = False                  ' 白黒印刷＝しない
        .Zoom = 80                              ' 印刷倍率
        .PrintErrors = xlPrintErrorsDisplayed   ' エラー表示の印刷＝見たまま印刷
        ' 以下、余白設定
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        ' 以下、フッター設定
        .RightFooter = "&""Century""&9&P"
    End With
End Sub
 
Private Sub publish_procedure()
    Dim PathPdf As String
    PathPdf = spread_fd & "\" & Replace(ActiveWorkbook.Name, ".xlsx", ".pdf")
    ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PathPdf, _
        Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
End Sub

Private Sub cover_procedure()
' 表紙設定
    Dim objFSO As Object, wb_obj As Object
    Dim wb_cover As Workbook, wb_crs As Workbook, wb_rd As Workbook
    Dim ws_cover As Worksheet
    Dim cover_fd As String, crs_fd As String, crs_fn As String, rd_fd As String, rd_fn As String
    Dim s_rw As Long, s_cnt As Long
    Dim tBox_Ctrl As Shape
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    cover_fd = file_path & "\5_INI"
    For Each wb_obj In objFSO.getfolder(cover_fd).Files
        If Right(wb_obj.Name, 8) = "cov.xlsx" Then
            Set wb_cover = Workbooks.Open(cover_fd & "\" & wb_obj.Name)
            Set ws_cover = wb_cover.Worksheets(1)
            ws_cover.Name = "表紙"
            ws_cover.Move before:=ws_spread0
            Application.DisplayAlerts = False
            wb_cover.Activate
            wb_cover.Close
            Application.DisplayAlerts = True

            Set ws_cover = wb_spread.Worksheets(1)
            
            For Each tBox_Ctrl In ws_cover.Shapes
                If tBox_Ctrl.Type = 17 Then
                    Select Case tBox_Ctrl.TextFrame.Characters.Text
                        Case "タイトル1"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 3 Then
                                If tn_tab = True And ws_mainmenu.Cells(3, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value & " 単純集計表"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(3, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value & " 集計表"
                                End If
                            ElseIf ws_mainmenu.Cells(6, 32).End(xlUp).Row > 3 Then
                                tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(3, 32).Value
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "タイトル1" Then
                                tBox_Ctrl.Delete
                            End If
                        Case "タイトル2"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 4 Then
                                If tn_tab = True And ws_mainmenu.Cells(4, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value & " 単純集計表"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(4, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value & " 集計表"
                                End If
                            ElseIf ws_mainmenu.Cells(6, 32).End(xlUp).Row > 4 Then
                                tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(4, 32).Value
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "タイトル2" Then
                                tBox_Ctrl.Delete
                            End If
                        Case "タイトル3"
                            If ws_mainmenu.Cells(6, 32).End(xlUp).Row = 5 Then
                                If tn_tab = True And ws_mainmenu.Cells(5, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(5, 32).Value & " 単純集計表"
                                ElseIf tn_tab = False And ws_mainmenu.Cells(5, 32).Value <> "" Then
                                    tBox_Ctrl.TextFrame.Characters.Text = ws_mainmenu.Cells(5, 32).Value & " 集計表"
                                End If
                            End If
                            If tBox_Ctrl.TextFrame.Characters.Text = "タイトル3" Then
                                tBox_Ctrl.Delete
                            End If
' 表紙に件数あるの何か違和感あるので、ちょっと保留… - 2018/6/22
'                        Case "集計対象件数："
'                            crs_fd = file_path & "\3_FD"
'                            crs_fn = Replace(spread_fn, "_集計表", "")
'                            Set wb_crs = Workbooks.Open(crs_fd & "\" & crs_fn)
'                            If WorksheetFunction.CountA(wb_crs.Worksheets(1).Range("Q:R")) = 5 Then
'                                rd_fn = wb_crs.Worksheets(1).Cells(2, 4).Value
'                                rd_fd = file_path & "\1_DATA"
'                                Set wb_rd = Workbooks.Open(rd_fd & "\" & rd_fn)
'                                s_rw = 6
'                                Do
'                                    If Cells(s_rw, 1).Offset(1, 0).Value <> "" And Cells(s_rw, 1).Value <> Cells(s_rw, 1).Offset(1, 0).Value Then
'                                        s_cnt = s_cnt + 1
'                                    ElseIf Cells(s_rw, 1).Offset(1, 0).Value = "" Then
'                                        Exit Do
'                                    End If
'                                    s_rw = s_rw + 1
'                                Loop
'                                tBox_Ctrl.TextFrame.Characters.Text = tBox_Ctrl.TextFrame.Characters.Text & s_cnt & " 件"
'                            End If
'                            wb_crs.Close
'                            wb_rd.Close
                        Case Else
                    End Select
                    ActiveWindow.View = xlNormalView
                    ActiveWindow.DisplayGridlines = False
                    ActiveWindow.Zoom = 80
                End If
            Next
            With ws_cover.Range("AB49")
                With .Font
                    .Name = "Arial"
                    .Color = RGB(255, 255, 255)
                End With
                .Value = "ACROSS Multiple Cross-tabulation System in" & Str(Year(Now)) & Space(1)
                .HorizontalAlignment = xlRight
                .FontSize = 8
            End With
        End If
    Next wb_obj
    
    Set wb_obj = Nothing
    Set ws_cover = Nothing
    Set wb_cover = Nothing
    Set wb_crs = Nothing
End Sub

Private Sub CoverMark_procedure()
    Dim ws_cover As Worksheet
    Dim tBox_Ctrl As Shape
    Set ws_cover = wb_print.Worksheets(1)
    For Each tBox_Ctrl In ws_cover.Shapes
        If tBox_Ctrl.Type = 17 And tBox_Ctrl.TextFrame.Characters.Text = "出力タイプ" Then
            Select Case wb_print.Worksheets(3).Name
                Case "Ｎ％表"
                    tBox_Ctrl.Delete
                Case "Ｎ表"
                    tBox_Ctrl.TextFrame.Characters.Text = "件数表"
                Case "％表"
                    tBox_Ctrl.TextFrame.Characters.Text = "構成比表"
                Case Else
            End Select
        End If
    Next
    Set ws_cover = Nothing
End Sub

