Attribute VB_Name = "Module04"
Option Explicit

' 各加工エリア列番号（定数）
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

' 各加工エリア列番号（定数）
Private Enum row_data

    START_ROW = 6
    
    ' セレクトフラグ判定用
    QCODE_P_CA = 6
    QCODE_P_ROW = 5
    QCODE_P_COLUMN = 4

End Enum

' 各加工エリア列番号（定数）
Private Enum ROW_INDATA

    START_ROW_INDATA = 7
    START_SUTATUSBER = 5
    
    START_TABLE_DATA = 6

End Enum

Public Sub processing_indata()
'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2017.05.18  '
' 入力データ加工プロシージャメインルーチン                                                         '
'--------------------------------------------------------------------------------------------------'
    Dim order_count As Long         ' 加工処理順番判定用変数
    Dim order_max As Long           ' 加工処理回数取得用変数
    
    Dim wb_process As Workbook      ' 加工指示ワークブック格納用オブジェクト
    Dim ws_menu As Worksheet        ' 加工指示メインメニュー格納用オブジェクト
    Dim ws_process As Worksheet     ' 加工指示作業ワークシート格納用オブジェクト
    
    Dim work_process As Long        ' 作業加工情報格納用変数
    
    Dim wsp_indata As Worksheet      ' 入力データワークシート格納用オブジェクト
    Dim indata_maxrow As Long       ' 入力データ最大列数格納用変数
    
    Dim indata_maxcolumn As Long    ' 入力データ最大行数格納用変数
    Dim indata_count As Long        ' 入力データヘッダ情報カウント用変数
    Dim hedder_flg As Boolean       ' ヘッダ情報フラグ（*加工後）
    Dim hedder_address As String    ' ヘッダ位置アドレス格納用変数
    
    Dim statusBar_text As String    ' ステータスバーコメント格納用変数
    
    Dim connect_data As String      ' 複数セレクト条件判定用変数

    Dim filename_work As String     ' 加工指示ファイル名格納用変数
    
    Dim error_tb As Workbook        ' エラー出力用ワークブックオブジェクト
    Dim error_ts As Worksheet       ' エラー出力用ワークシートオブジェクト
    
    Dim indata_logname As String    ' ログ出力用インデータネーム
    
'    Dim now_data As Databar         ' 日時格納用変数

    Dim check_wb As Workbook        ' オープンチェック用ワークブックオブジェクト
    Dim check_flg As Boolean        ' オープンチェック結果格納用フラグ
    Dim check_name As String        ' オープンチェック用ファイル名格納用変数
    
    Call Indata_Open
    Call Setup_Hold
    Call Filepath_Get
    Call Setup_Check
    
    indata_logname = wb_indata.Name
    
    ' 画面への表示をオフにする
    Application.ScreenUpdating = False
    
    ' メッセージを非表示にする
    Application.DisplayAlerts = False
    
    ' 入力データをワークシートとして確保
    Set wsp_indata = wb_indata.Worksheets(1)
    
    ' 入力データ最大行
    indata_maxcolumn = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' ヘッダ情報を格納
    For indata_count = 1 To indata_maxcolumn
        ' ヘッダ情報に*加工後があるかを判定
        If wsp_indata.Cells(1, indata_count) = "*加工後" Then
            hedder_flg = True
        End If
    Next
    
    ' ヘッダに*加工後が含まれていなかった場合
    If hedder_flg = False Then
        
        ' 先頭のアドレスを取得
        hedder_flg = Hedder_Create(wsp_indata, "*加工後", wsp_indata.Cells(1, indata_count).Address)
        
'        ' コメントを入力データヘッダに代入、下線を引く
'        wsp_indata.Range(hedder_address).Value = "*加工後"
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'
'        ' ヘッダアドレスを変更
'        hedder_address = wsp_indata.Range(hedder_address).Offset(1).Resize(3).Address
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
'
'        ' 全体に罫線を引く
'        hedder_address = wsp_indata.Range(hedder_address).Offset(3).Resize(2).Address
'        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlDash
'        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).Weight = xlHairline
'
'        ' 列に色をつける
'        wsp_indata.Columns(indata_count).Interior.Color = RGB(255, 255, 0)
        
    End If
    
    ' 入力データ最大列数を格納
    indata_maxrow = wsp_indata.Cells(Rows.Count, 1).End(xlUp).Row
    
' ファイル名取得処理
step00:

    ChDrive file_path & "\3_FD"
    ChDir file_path & "\3_FD"
    
    filename_work = Application.GetOpenFilename("加工指示ファイル,*.xlsm", , "加工指示ファイルを開く")
    If filename_work = "False" Then
        ' キャンセルボタンの処理
        End
    ElseIf filename_work = "" Then
        MsgBox "［加工指示ファイル］を選択してください。", vbExclamation, "MCS 2017 - processing_indata"
        GoTo step00
    ElseIf InStr(filename_work, "加工指示") = 0 Then
        MsgBox "［加工指示ファイル］を選択してください。", vbExclamation, "MCS 2017 - processing_indata"
        GoTo step00
    End If
    
    ' ファイル名のみを抽出
    check_name = Mid(filename_work, Len(file_path & "\3_FD\"))
    
    ' ファイルを開いているかをチェック
    For Each check_wb In Workbooks
        If check_wb.Name = check_name Then
            check_flg = True
        End If
    Next check_wb
    
    ' ブックを開いているか判定
    If check_flg = True Then
        Set ws_menu = wb_process.Worksheets("メインメニュー")
        
    Else
        Set wb_process = Workbooks.Open(filename_work)
        Set ws_menu = wb_process.Worksheets("メインメニュー")
    
    End If
    
    ' 加工回数を取得
    order_max = ws_menu.Range("AE29").Value
    
    ' 20180614
    ' 加工ログ内容出力用ファイルを確認
    Set error_tb = Workbooks.Add
    Set error_ts = error_tb.Worksheets(1)
    
    ' 加工出力のヘッダを作成
    error_ts.Range("A1").Value = "SEQ"
    error_ts.Range("B1").Value = "加工内容"
    error_ts.Range("C1").Value = "QCode1"
    error_ts.Range("D1").Value = "QCode2"
    error_ts.Range("E1").Value = "処理内容"
    
    ' 幅調整
    error_ts.Columns(1).ColumnWidth = 6
    error_ts.Columns(2).ColumnWidth = 20
    error_ts.Columns(3).ColumnWidth = 10
    error_ts.Columns(4).ColumnWidth = 10
    error_ts.Columns(5).ColumnWidth = 70
    
    ' その他調整
    error_ts.Name = "加工内容一覧"
    
    
    ' 優先順に合わせて加工処理を行う
    For order_count = 1 To order_max
        
        statusBar_text = "入力データ加工作業中(" & Format(order_count) & "/" & order_max & ")"
        Application.StatusBar = statusBar_text
        
        ' 行う作業内容を取得
        work_process = wb_process.Worksheets("メインメニュー").Cells(31 + (order_count - 1), 31).Value
    
        ' 作業内容に合わせて各加工処理をコールする
        Select Case work_process
            
            ' 逆セット処理
            Case 1
            
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("逆セット処理")
                Call Processing_Setreverse(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' ○＋◎処理
            Case 2
                
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("○+◎処理")
                Call Processing_Complementarity(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' 排他的処理
            Case 3
                
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("排他的処理")
                Call Processing_Exclusive(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            ' セレクトフラグ処理
            Case 4
            
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("セレクトフラグ処理")
                'Call Processing_Selectflg(ws_process, wsp_indata, indata_maxrow, statusBar_text)
                Call Processing_Selectflg3(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            ' カテゴライズ処理
            Case 5
            
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("カテゴライズ処理")
                Call Processing_Categorize2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts, error_tb)
                
            Case 6
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("データクリア処理")
                Call data_clear1(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            Case 15
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("リミットマルチ加工処理")
                Call Processing_Limitmulti_2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
            
            Case 16
                ' 作業シートを固定し処理を行う、設定画面情報から処理を行う為ws_processは未使用
                Set ws_process = wb_process.Worksheets("データクリア処理")
                Call data_clear2(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)

            Case 17
                ' 作業シートを固定し処理を行う
                Set ws_process = wb_process.Worksheets("増幅加工処理")
                Call amplification_data(ws_process, wsp_indata, indata_maxrow, statusBar_text, error_ts)
                
            Case 18
                ' ファイル名を変更し保存
                
                ' 加工処理が行われなかった場合
                If wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Value = "*加工後" Then
                
                    ' 加工後行を削除する
                    wsp_indata.Columns(wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column).Delete
                
                End If
                
                wb_indata.SaveAs _
                Filename:=file_path & "\1_DATA\" & ThisWorkbook.Worksheets("メインメニュー").Range("H3").Value & "OT.xlsx"
                
            Case Else
    
        End Select
    
    Next order_count
        
'    Call Processing_Setreverse      ' 逆セット加工用プロシージャ
'    Call Processing_Exclusive       ' 排他的処理加工用プロシージャ
'    Call Processing_Categorize      ' カテゴライズ処理加工用プロシージャ
'    Call Processing_Complementarity ' ○＋◎処理加工用プロシージャ
    
    ' ステータスバーを初期化
    Application.StatusBar = False
    
    ' 画面への表示をオンにする
    Application.ScreenUpdating = True
    
    ' カテゴライズ用ログの作成
    'now_data = Now
    error_tb.SaveAs Filename:=file_path & "\4_LOG\" & Format(Now, "yyyymmddhhmmss") & _
    "_" & Mid(wb_process.Name, 1, Len(wb_process.Name) - 5) & "(加工ログ).xlsx"
    error_tb.Close
    
    ' module04 エラーもしくはログ出力ファイルを閉じる
    'Close #30
    
    wb_process.Close SaveChanges:=False
    
    ' 入力ファイルをクローズする
    wb_indata.Close
    
    ' 0byteファイルを削除
    Call Finishing_Mcs2017
    Call Starting_Mcs2017
    
    ' システムログの出力
    ' 2020.6.4 - 追加
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
    Print #1, Format(Now, "yyyy/mm/dd hh:mm:ss") & " - 入力データの加工処理：使用ファイル［" & indata_logname & " ｜ " & Dir(filename_work) & "］"
    Close #1
    
    
    ' メッセージを非表示にする
    Application.DisplayAlerts = True
    

    
    MsgBox "入力データの加工が完了しました。", , "MCS2017"

End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.04.19  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2018.06.14　'
' 逆セット加工用プロシージャ                                                                       '
' 引数１ WorkSheet型 逆セット処理指示シート                                                        '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Setreverse(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim qcode1_dataflg As Boolean   ' 比較設問(子) 判定フラグ
    Dim qcode2_dataflg As Boolean   ' 逆セット対象設問(親)判定フラグ

    Dim ma_count As Long            ' MA回答内容確認用カウント変数
    
    Dim force_setflg As Boolean     ' 強制逆セットフラグ格納用変数
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　逆セット加工処理中(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
        
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
        
        ' 対象設問の行番号を取得
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
        qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2))
        
        ' InputWordの取得
        input_word = ws_process.Cells(process_count, QCODE2_DATA1).Value
        
        ' 各種フラグを初期化
        process_flg = False
        force_setflg = False
        
        ' 強制判定を行う ※処理を行わない場合は"FALSE"
        If ws_process.Cells(process_count, QCODE2_DATA2).Value <> "" Then
            force_setflg = True
        End If
        
        ' 処理判定を行う ※処理を行わない場合は"FALSE"
        If ws_process.Cells(process_count, SKIP_FLG).Value = "" Then
            process_flg = True
        End If
        
        ' Input_Wordが選択肢範囲外の時（マルチアンサー）
        If Val(input_word) > q_data(qcode2_row).ct_count And q_data(qcode2_row).q_format = "M" Then
            Call print_log("逆セット処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
            "※書き込み先の" & q_data(qcode2_row).q_code & "に「" & input_word & "」は存在しないため処理を行わずに終了", ws_logs)
            process_flg = False
        End If
        
        ' Input_Wordが選択肢範囲外の時（リミットマルチ）
        If Val(input_word) > q_data(qcode2_row).ct_count And Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
            Call print_log("逆セット処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
            "※書き込み先の" & q_data(qcode2_row).q_code & "に「" & input_word & "」は存在しないため処理を行わずに終了", ws_logs)
            process_flg = False
        End If
        
        ' 処理を行う（処理判定が有効かつ対応する指示が全て行われている時）
        If process_flg = True And qcode1_row <> 0 And qcode2_row <> 0 And input_word <> "" Then
        
            ' クリア処理が行われない時　（親が複数回答可）
            If force_setflg = True Then
            
                Call print_log("逆セット処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "に回答がある場合、" & q_data(qcode2_row).q_code & "に強制的に「" _
                & input_word & "」を入力", ws_logs)
            
            ' クリア処理が行われない時　（親が複数回答可）
            ElseIf q_data(qcode1_row).q_format = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
            
                Call print_log("逆セット処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "に回答があり、" & q_data(qcode2_row).q_code & "が無回答の時に「" _
                & input_word & "」を入力", ws_logs)
            
            ' クリア処理が行われるとき　（親が単一回答）
            Else
                Call print_log("逆セット処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, _
                q_data(qcode1_row).q_code & "に回答があり、" & q_data(qcode2_row).q_code & "が無回答の時に「" _
                & input_word & "」を入力（" & q_data(qcode2_row).q_code & "に「" & input_word & "」以外の回答があれば" _
                & q_data(qcode1_row).q_code & "をクリア", ws_logs)
            End If
        
        
            ' 入力データ全てを判定する
            For indata_count = START_ROW_INDATA To indata_maxrow
            
                ' 各種フラグを初期化
                qcode1_dataflg = False
                qcode2_dataflg = False
                
                ' 子の設問がMAもしくはLMの時
                If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
                
                    ' 複数回答の内容を確認し、回答があればqcode1_dataflgを有効にする
                    For ma_count = 1 To q_data(qcode1_row).ct_count
                
                        ' いずれかのカテゴリーに回答がある場合
                        If Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ma_count - 1))) <> "" Then
                            qcode1_dataflg = True
                        End If
                
                    Next ma_count
                
                Else
                    
                    ' 入力データの内容を判定
                    If Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value) <> "" Then
                
                        ' 処理を行う
                        qcode1_dataflg = True
                        
                    Else
                        
                        ' 処理を行わない
                        qcode1_dataflg = False
                        
                    End If
                
                End If
                
                ' Qcode1が処理可能の場合
                If qcode1_dataflg = True Then
                
                    ' 親の設問がMAもしくはLMの時
                    If Mid(q_data(qcode2_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
                    
                        ' 20170519 0ct 対応
                        ' 0カテゴリーフラグが有効の時
                        'If q_data(qcode2_row).ct_0flg = True Then
                        
                            ' 逆セット処理を行う
                        '    wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + Val(input_word)) = 1
                        
                        ' 通常処理
                        'Else
                        
                            ' 逆セット処理を行う
                            wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + (Val(input_word) - 1)) = 1
                        
                        'End If
                    
                    Else
                                            
                        ' 親の回答内容を判定（親が無回答の時）
                        If Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) = "" Or _
                        ws_process.Cells(process_count, QCODE2_DATA2).Value <> "" Or force_setflg = True Then
                            
                            ' 逆セット処理を行う
                            wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value = input_word
                        
                        ' 親の回答内容を判定（親に異なる回答がある時）
                        ElseIf Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) <> Trim(input_word) Then
                        
                            ' 子の設問がMAもしくはLMの時
                            If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
                                    
                                DoEvents
                                    
                                ' 子の回答を全て無回答処理
                                For ma_count = 1 To q_data(qcode1_row).ct_count
                
                                    wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ma_count - 1)) = ""
                
                                Next ma_count
                            
                            ' 子の回答が自由記述の時
                            ElseIf q_data(qcode1_row).q_format = "F" Or q_data(qcode1_row).q_format = "O" Then
                                
                                ' クリアは行わない
                                
                            Else
                            
                                ' 子の回答を無回答処理
                                wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value = ""
                        
                            End If
                        
                        ' 親の回答内容を判定（親の回答が同一の時）
                        ElseIf Trim(wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Value) <> input_word Then
                            
                            ' 処理を行わない
                            
                        End If
                        
                    End If
                    
                End If
                
            Next indata_count
        
        End If
        
    Next process_count

End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.04.19  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2018.06.14　'
' 排他的処理加工用プロシージャ                                                                     '
' 引数１ WorkSheet型 排他的処理指示シート                                                          '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Exclusive(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
'    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
'    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim qcode1_dataflg As Boolean   ' 比較設問(子) 判定フラグ
'    Dim qcode2_dataflg As Boolean   ' 逆セット対象設問(親)判定フラグ
    Dim work_maxcol As Long         ' 指示回答終端位置

    Dim ma_count As Long            ' MA回答内容確認用カウント変数
    
    Dim exclusive_ct As Long        ' 排他的処理対象番号格納用変数
    
    Dim str_ct() As Variant         ' カテゴリー比較用配列（入力データ）
    Dim str_maxcount As Long        ' 回答数格納用変数
    Dim str_min As Long             ' 最小値取得用変数
    Dim str_address As String       ' MA先頭アドレス格納用文字列変数
    Dim str_target As Long          ' 配列内比較用カウント変数
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　排他的処理中(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' QCODEを検索し設定画面から列番号を取得
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
    
        ' 指示がMAかLMの時
        If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Then
            
            ' 排他的カテゴリーの指示が行われているときかつスキップ指示がない時
            If ws_process.Cells(process_count, QCODE1_DATA1).Value <> "" _
            And ws_process.Cells(process_count, SKIP_FLG).Value = "" Then
            
                ' 排他的処理カテゴリー番号を取得
                exclusive_ct = ws_process.Cells(process_count, QCODE1_DATA1).Value
                
                ' カテゴリー番号が設問の範囲内かどうかを判定（範囲内）
                If q_data(qcode1_row).ct_count >= exclusive_ct And exclusive_ct <> 0 Then
                    
                    Call print_log("排他的処理", q_data(qcode1_row).q_code, "", q_data(qcode1_row).q_code & _
                    "の「" & exclusive_ct & "」と他の回答が混在している時、「" & exclusive_ct & "」をクリア", ws_logs)
                    
                    ' 入力データに処理を行う
                    For indata_count = START_ROW_INDATA To indata_maxrow
                    
                    
                        ' 対象のカテゴリーが有効かを判定
                        If wsp_indata.Cells(indata_count, _
                        q_data(qcode1_row).data_column + (exclusive_ct - 1)).Value = 1 Then
                   
                            ' MA範囲を全て配列に格納する
                            str_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            str_ct = wsp_indata.Range(str_address).Resize(, q_data(qcode1_row).ct_count)
                            ' ※ str_ct( 1 to 1 , 1 to q_data(qcode1_row).ct_count )
                    
                            ' 配列への回答数を取得
                            str_maxcount = Application.WorksheetFunction.Sum(str_ct)
                            str_min = Application.WorksheetFunction.Min(str_ct)
                            
                            ' 対象のカテゴリー以外にも回答がある場合
                            If str_maxcount > 1 Then
                                
                                ' 無回答へ変更する
                                wsp_indata.Cells(indata_count, _
                                q_data(qcode1_row).data_column + (exclusive_ct - 1)).Value = ""
                        
                            End If
                    
                        End If
                    
                    Next indata_count
                
                ' カテゴリー番号が設問の範囲内かどうかを判定（範囲外）
                Else
                    Call print_log("排他的処理", q_data(qcode1_row).q_code, "", "※" & q_data(qcode1_row).q_code & _
                    "の選択肢「" & exclusive_ct & "」は存在していないため、処理を終了をしました。", ws_logs)
                End If
                
            End If
            
        End If
        
    Next process_count
    
End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.04.20  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.04.21　'
' ○＋◎処理加工用プロシージャ                                                                     '
' 引数１ WorkSheet型 ○＋◎処理指示シート                                                          '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Complementarity(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim work_maxcol As Long         ' 指示回答終端位置
    
    Dim str_ct1 As Variant          ' カテゴリー比較用配列（◎）
    Dim str_ct2 As Variant          ' カテゴリー比較用配列（○）
    Dim str_ct3 As Variant          ' カテゴリー比較用配列（○+◎）
    
    Dim str_maxcount As Long        ' 回答数格納用変数
    Dim str_address As String       ' MA先頭アドレス格納用文字列変数
    Dim str_target As Long          ' 配列内比較用カウント変数
    
    Dim ct_count As Long            ' 配列内容カウント用変数
    
    Dim ct1_count As Long           ' ◎回答数格納用変数
    Dim ct2_count As Long           ' ○回答数格納用変数
    Dim ct3_count As Long           ' ○+◎回答数格納用変数
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
        
        'Application.StatusBar = statusBar_text & _
        '"　○+◎処理中(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' スキップフラグがたっていない時かつ各QCODEが全て記入されている時
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE2) <> "" Then
        
        ' QCODEを検索し設定画面から列番号を取得
        qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
        qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2))
        
            ' ◎指示がMAかLMかSの時
            If Mid(q_data(qcode1_row).q_format, 1, 1) = "M" Or _
            Mid(q_data(qcode1_row).q_format, 1, 1) = "L" Or _
            Mid(q_data(qcode1_row).q_format, 1, 1) = "S" Then
                
                ' ○指示がMAかLMの時
                If Mid(q_data(qcode2_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode2_row).q_format, 1, 1) = "L" Then
        
                    ' 処理内容を確認 ◎→○ and ○→◎
                    If ws_process.Cells(process_count, QCODE2_DATA1) = "" Then
                        Call print_log("○＋◎処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & q_data(qcode1_row).q_code & _
                        "（◎）の回答を" & q_data(qcode2_row).q_code & "（○）に追加し、◎が無回答かつ○の回答数が◎の有効回答数以内の時、○の回答内容を◎に追加", ws_logs)
                    ' 処理内容を確認 ◎→○
                    Else
                        Call print_log("○＋◎処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & q_data(qcode1_row).q_code & _
                        "（◎）の回答を" & q_data(qcode2_row).q_code & "（○）に追加", ws_logs)
                    End If
                    
                    ' 入力データに処理を行う
                    For indata_count = START_ROW_INDATA To indata_maxrow
        
                        ' 同一サイズの配列を再定義
                        ReDim str_ct1(1, q_data(qcode1_row).ct_count)
                        ReDim str_ct2(1, q_data(qcode2_row).ct_count)
                        ReDim str_ct3(q_data(qcode1_row).ct_count)
        
                        ' ◎指示がSの時
                        If q_data(qcode1_row).q_format = "S" Then
                        
                            ' 回答があった時
                            If wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <> "" Then
                            
                                ' カテゴリー位置に1を代入
                                str_ct1(1, wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column)) = 1
                        
                            End If
                        
                        ' ◎指示がMAかLMの時
                        Else
                        
                            ' MA範囲を全て配列に格納する
                            str_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            str_ct1 = wsp_indata.Range(str_address).Resize(, q_data(qcode1_row).ct_count)
                        
                        End If
                        
                        ' ○指示を全て配列に格納する
                        str_address = wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column).Address
                        str_ct2 = wsp_indata.Range(str_address).Resize(, q_data(qcode2_row).ct_count)
                        
                        ' 最大カテゴリー数を取得（○の方がカテゴリーが大きいためstr_ct2を使用）
                        ' 20180618 小さいカテゴリー数に合わせ配列を用意
                        If q_data(qcode1_row).ct_count = q_data(qcode2_row).ct_count Or _
                        q_data(qcode1_row).ct_count > q_data(qcode2_row).ct_count Then
                            str_maxcount = q_data(qcode2_row).ct_count
                        ElseIf q_data(qcode1_row).ct_count < q_data(qcode2_row).ct_count Then
                            str_maxcount = q_data(qcode1_row).ct_count
                        End If
                        

                        
                        ' ◎(str_ct1)の回答状況と○(str_ct2)の回答状況をあわせstr_ct3を作成する
                        For ct_count = 1 To str_maxcount
                        
                            'str_ct3(ct_count) = str_ct1(1, ct_count) Or str_ct2(1, ct_count)
                            ' ○もしくは◎で回答がある場合
                            If str_ct1(1, ct_count) > 0 Or str_ct2(1, ct_count) > 0 Then
                            
                                ' str_ct3にあわせたデータを作成する
                                ' 20180618 加工段階では回答の内容問わず1立てのみ
                                str_ct3(ct_count) = 1
                            
                            Else
'                               str_ct3(ct_count) = ""
                            End If
                        
                        Next ct_count
                        
                        ' 各配列の回答数を格納
                        ct1_count = Application.WorksheetFunction.Sum(str_ct1)
                        ct2_count = Application.WorksheetFunction.Sum(str_ct2)
                        ct3_count = Application.WorksheetFunction.Sum(str_ct3)
                        
                        DoEvents
                        
                        ' 合算して回答数が増えている場合
                        If ct2_count < ct3_count Then
                            
'                            wsp_indata.Range(str_address).Resize(, q_data(qcode2_row).ct_count) = str_ct3

                            ' str_ct3の内容を○(str_ct2)へ上書きする
                            For ct_count = 1 To str_maxcount
                            
                                ' str_ct3にあわせたデータを作成する
                                wsp_indata.Cells(indata_count, q_data(qcode2_row).data_column + (ct_count - 1)) _
                                = str_ct3(ct_count)
 
                            Next ct_count
                            
                            ' カテゴリー数も合わせて変更
                            ' 20180618 合算範囲を変更したため削除
                            'ct2_count = ct3_count
                            
                        End If
                        
                        ' セットを行う時
                        If ws_process.Cells(indata_count, QCODE2_DATA1).Value = "" Then
                        
                            ' ◎がSAかつ◎の回答数が0、○の回答数が1の時
                            If q_data(qcode1_row).q_format = "S" And ct1_count = 0 And ct2_count = 1 Then
                        
                                ' str_ct2の内容を◎(str_ct1)へ上書きする
                                For ct_count = 1 To str_maxcount
                            
                                    ' ○の回答内容と同じカテゴリーを有効にする
                                    If str_ct2(1, ct_count) > 0 Then
                                
                                        ' ○を◎へ追加入力
                                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value = ct_count
                                        '
                                        Exit For
                                
                                    End If
                            
                                Next ct_count
                        
                            ' ◎が無回答かつ○の回答数が◎のループカウント以下の時
                            ElseIf ct1_count = 0 And q_data(qcode1_row).ct_loop >= ct2_count Then
                        
                                ' 最大カテゴリー数を取得（◎のカテゴリー数に合わせる為str_ct1を使用）
                                str_maxcount = q_data(qcode2_row).ct_count
                        
                                ' str_ct2の内容を○(str_ct1)へ上書きする
                                For ct_count = 1 To str_maxcount
                            
                                    ' str_ct3にあわせたデータを作成する
                                    wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column + (ct_count - 1)) _
                                    = str_ct2(1, ct_count)
 
                                Next ct_count
                        
                            End If
                        
                        End If
                        
                    Next indata_count
        
                ' ○指示が複数回答ではない時
                Else
                
                    ' エラーコメントの出力
                    Call print_log("○＋◎処理", q_data(qcode1_row).q_code, q_data(qcode2_row).q_code, "" & _
                    q_data(qcode2_row).q_code & "（○）の形式が複数回答設問ではありません", ws_logs)
                
                End If
                
            End If

        End If
    
    Next process_count


End Sub

'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2017.05.08  '
' セレクトフラグ加工用プロシージャ                                                                 '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
    Dim qcode3_row As Long          ' エントリーエリア終端格納用変数
    
    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim work_maxcol As Long         ' 指示回答終端位置
    
    Dim str_ct1 As Variant          ' カテゴリー比較用配列（◎）
    Dim str_ct2 As Variant          ' カテゴリー比較用配列（○）
    Dim str_ct3 As Variant          ' カテゴリー比較用配列（○+◎）
    
    Dim str_maxcount As Long        ' 回答数格納用変数
    Dim str_address As String       ' MA先頭アドレス格納用文字列変数
    Dim str_target As Long          ' 配列内比較用カウント変数
    
    Dim ct_count As Long            ' 配列内容カウント用変数
    
'    Dim ct1_count As Long           ' ◎回答数格納用変数
'    Dim ct2_count As Long           ' ○回答数格納用変数
'    Dim ct3_count As Long           ' ○+◎回答数格納用変数
    
    Dim work1_flg As Boolean        ' 第一条件格納用フラグ
    Dim work2_flg As Boolean        ' 第二条件格納用フラグ
    
    Dim column_max As Long          ' AND、OR条件終端番号格納用変数
    Dim and_column() As Long        ' 複数条件対応用動的配列
    Dim and_data() As Variant       ' 複数条件情報格納用変数
    Dim and_count As Long           ' AND、OR条件
    Dim and_target As Long          ' 対象列番号格納用変数
    
    Dim processing_flg As Boolean   ' Function戻り値格納用変数
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 複数条件の終端位置を取得
    column_max = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' アドレス位置格納用配列を再定義
    ReDim and_column(300)
    
    ' セレクト条件の初期位置を設定
    and_count = 1
    and_column(1) = SF_START
    
    ' アドレス位置を取得
    For and_target = QCODE1_DATA6 To column_max
                    
        ' 条件の行を判定
        If ws_process.Cells(START_ROW - 1, and_target).Value = "接続詞" & vbLf & "（複数条件）" Then
        
            ' 書き込み位置を変更
            and_count = and_count + 1
            
            ' カラム情報を配列に格納
            and_column(and_count) = and_target
            
        End If
        
    Next
    
    ' 配列を再定義（値は保持する）
    ReDim Preserve and_column(and_count)
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　セレクト加工処理中(" & Format(process_count - START_SUTATUSBER) & _
        '"/" & Format(process_max - START_SUTATUSBER) & ")"

        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' スキップフラグがたっていない時かつ各QCODEが全て記入されている時
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA3) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA4) <> "" Then
        
            ' QCODEをマッチングし情報を列番号を格納
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1))
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4))
            qcode3_row = Qcode_Match("*加工後")
        
            ' 既に出力エリアが用意されている時（かつ加工後よりも後に出力エリアが設定されている時
            If q_data(qcode2_row).data_column <> 0 And _
            q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' 通常処理
                If q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' EntryAreaに書き込む指示の場合
                Else
                                        
                    'Print #30, "加工先がEntryAreaに設定されているため処理を行いませんでした、確認をお願いします。"
            
                End If
            
            ' まだエリアが用意されていない時
            Else
        
                ' ヘッダを作成する
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' 新しく設定したエリアのq_dataにカラムとして設定する
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' 最小値、最大値を入力データのヘッダに格納
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' 加工指示の内容を列ごと取得
            and_data = ws_process.Range(ws_process.Cells(process_count, 1).Address).Resize(, column_max - 1)
            'ReDim Preserve and_data(column_max - 1)
            
            ' セレクトフラグの条件に合わせて回答内容を判定する
            Call salectflg_decision(wsp_indata, and_data, and_column, indata_maxrow, qcode2_row)
            'Call data_clear(wsp_indata, and_)
        
        
        ' 配列終端まで判定
        'For column_count = 1 To UBound(and_column)
        'Next
    
        End If
    
    Next process_count
    
End Sub


'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.04.21  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.02　'
' カテゴライズ処理加工用プロシージャ                                                               '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Categorize(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim str_address() As String     ' 処理アドレス格納用配列
    Dim str_qcode() As Long         ' 処理QCODE格納用配列
    Dim target_coderow As Long      ' 対象QCODE列番号一時格納用変数
    
    Dim column_count As Long        ' 処理列番号格納用変数
    Dim column_end As Long          ' 列終端番号格納用変数
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数

    Dim str_ctdata As Variant       ' カテゴライズテーブル情報格納用配列変数
    Dim str_ctdata_c As Range       ' 一時保存用変数
    
    Dim writing_column As Long      ' 書き込み位置格納用変数
    Dim target_data As Double       ' 入力データ判定用変数
    Dim categorize_count As Long    ' 判定カウント用変数
    Dim table_max As Long           ' 終端テーブル番号格納用変数
    Dim ma_count As Long            ' MAカテゴリー情報を取得
    
    ' 入力データMA操作用変数群
    Dim str_madata As Variant       ' MA回答データ格納用変数
    Dim ma_address As String        ' MA回答データアドレス格納用変数
    Dim maindata_count As Long      ' MAデータカウント変数
    
    Dim prosessing_flg As Boolean   ' ヘッダ加工判定用フラグ

    Dim pmax_number As Double       ' テーブル最大値格納用変数
    Dim pmin_number As Double       ' テーブル最小値格納用変数

    Dim str_count As Long           ' str_ctdataカウント用変数

    ' 初期設定
    ReDim str_address(300)
    ReDim str_qcode(300)
    
    ' 画面への表示をオンにする
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "　カテゴライズ加工処理数計算中..."
    
    ' 画面への表示をオフにする
    'Application.ScreenUpdating = False
    
    ' 終端列番号の取得 20170502 START_ROW-1を5に変更
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column

    ' 処理件数を取得
    For column_count = 11 To column_end
    
        ' アスタリスクを見つけたとき数を数える
        ' ※アスタリスクと割当ＣＴが固定位置＆対象設問に記入あり＆Skipフラグ無し
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 1, column_count + 5).Value = "割当ＣＴ" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 5).Value = "" Then
            
            ' QCODEを取得
            target_coderow = Qcode_Match(ws_process.Cells(START_ROW - 2, column_count + 2).Value)
            
            ' 内容が数値の時
            If IsNumeric(target_coderow) Then
            
                ' 件数を対象件数を増やす
                process_max = process_max + 1
            
                ' QCODE列番号を配列に格納する
                str_qcode(process_max) = target_coderow
            
                ' アスタリスクの位置を配列に格納する
                str_address(process_max) = ws_process.Cells(START_ROW - 1, column_count).Address

            End If
            
        End If
    
    Next column_count
    
    ' 格納した件数に応じて配列を再定義（値は保持する）
    ReDim Preserve str_address(process_max)
    
    ' カテゴライズ処理を行う（配列に取り込んだ指示数分）
    For process_count = 1 To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　カテゴライズ加工処理中(" & Format(process_count) & "/" & Format(process_max) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' テーブル情報を配列に格納
        Set str_ctdata_c = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
        str_ctdata = str_ctdata_c.NumberFormatLocal
        
        ' 書き込み位置を取得
        writing_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        
        ' 入力データエリアにQCODEを設定
        ' wsp_indata.Cells(1, writing_column) = ws_process.Range(str_address(process_count)).Offset(-1, 3).Value
        
        ' ヘッダ情報を作成
        prosessing_flg = Hedder_Create(wsp_indata, ws_process.Range(str_address(process_count)) _
        .Offset(-1, 3).Value, wsp_indata.Cells(1, writing_column).Address)
        
        ' ＭＡ出力情報を取得
        ma_count = Val(ws_process.Range(str_address(process_count)).Offset(-1, 4))
        
        ' ＭＡ以外で出力を行う時
        If ma_count = 0 Then
            
            ' 初期値設定
            pmax_number = Val(str_ctdata(1, 6))
            pmin_number = Val(str_ctdata(1, 6))
            
            ' ヘッダー作成用
            For str_count = 1 To 300
            
                ' 最大値、最小値を取得
                If str_ctdata(str_count, 6) <> "" Then
                
                    ' 最大値
                    If pmax_number < Val(str_ctdata(str_count, 6)) Then
                        pmax_number = Val(str_ctdata(str_count, 6))
                    End If
                    
                    ' 最小値
                    If pmin_number > Val(str_ctdata(str_count, 6)) Then
                        pmin_number = Val(str_ctdata(str_count, 6))
                    End If
                
                Else
                    Exit For
                End If
            
            Next
            
            ' 最小値と最大値を書き込み
            wsp_indata.Cells(5, writing_column).Value = pmin_number
            wsp_indata.Cells(6, writing_column).Value = pmax_number
        
        End If
        
        ' 終端テーブル番号を取得
        table_max = UBound(str_ctdata)
        
        ' 入力データ全てにカテゴライズ処理を行う
        For indata_count = START_ROW_INDATA To indata_maxrow
            
            ' 元の回答区分によって処理をわける
            Select Case q_data(str_qcode(process_count)).q_format

                ' SAもしくは実数回答系列
                Case "S", "R", "H"
            
                    ' 入力値が存在する場合処理を行う
                    If wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value <> "" Then
                        
                        ' 入力データを取得
                        target_data = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
            
                        ' テーブル分の処理をまとめて行う
                        For categorize_count = 1 To table_max
                    
                            ' カテゴライズ指示が無い場合
                            If str_ctdata(categorize_count, 3) = "" Then
                                Exit For
                            End If
                    
                            ' アスタリスク箇所に情報が記入されていない時処理を行う
                            If str_ctdata(categorize_count, 1) = "" Then
                        
                                ' テーブルの範囲内の記入値の場合
                                If target_data >= str_ctdata(categorize_count, 3) And target_data <= _
                                str_ctdata(categorize_count, 5) Then
                            
                                    ' MA出力指示を判定(SA出力)
                                    If ma_count = 0 Then
                                        
                                        ' 割当CTを出力先に設定
                                        wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 6)
                                        
                                    ' MA出力指示を判定 (MA出力)
                                    Else
                                        
                                        ' 割当CTを出力先に設定
                                        wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 6) - 1)) = 1
                            
                                    End If
                        
                                ' テーブル範囲外の記入値の場合
                                Else
                                
                                    ' ※ログを出力する
                                
                                End If
                        
                            End If
            
                        Next categorize_count
                    End If
                
                ' MA
                Case "M"
                    
                    ' 入力データの先頭アドレスを取得
                    ma_address = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Address
                    
                    ' MA回答範囲を配列として取得
                    str_madata = wsp_indata.Range(ma_address).Resize(, (q_data(str_qcode(process_count)).ct_count))
                    'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
                                        
                    ' 回答が存在している場合
                    If WorksheetFunction.Sum(str_madata) <> 0 Then
                    
                        ' MA情報をカテゴライズ
                        For maindata_count = 1 To q_data(str_qcode(process_count)).ct_count
                        
                            ' 入力データが存在している場合（０以上の数値）
                            If str_madata(1, maindata_count) > 0 Then
            
                                ' テーブル分の処理をまとめて行う
                                For categorize_count = 1 To table_max
                    
                                    ' カテゴライズ指示が無い場合
                                    If str_ctdata(categorize_count, 3) = "" Then
                                        Exit For
                                    End If
                    
                                    ' アスタリスク箇所に情報が記入されていない時処理を行う
                                    If str_ctdata(categorize_count, 1) = "" Then
                        
                                        ' テーブルの範囲内の記入値の場合
                                        If maindata_count >= str_ctdata(categorize_count, 3) And maindata_count <= _
                                        str_ctdata(categorize_count, 5) Then
                            
                                            ' 割当CTを出力先に設定
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 6) - 1)) = 1
                        
                                        
                                        ' テーブルの範囲外の記入値の場合
                                        
                                            ' ※ログを出力する
                                        
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
' 作　成　者 村山 誠                                                           作成日  2017.04.21  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.02　'
' カテゴライズ処理加工用２プロシージャ                                                             '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
' 引数５ Workbook型  加工ログ出力ブック                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Categorize2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet, ByVal error_tb As Workbook)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim str_address() As String     ' 処理アドレス格納用配列
    Dim str_qcode() As Long         ' 処理QCODE格納用配列
    Dim str_outqcode() As String    ' 出力QCODE格納用配列
    Dim target_coderow As Long      ' 対象QCODE列番号一時格納用変数
    
    Dim column_count As Long        ' 処理列番号格納用変数
    Dim column_end As Long          ' 列終端番号格納用変数
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数

    Dim str_ctdata As Variant       ' カテゴライズテーブル情報格納用配列変数
    Dim str_ctpro As String         ' カテゴライズテーブル書式変更用変数
    
    Dim writing_column As Long      ' 書き込み位置格納用変数
    Dim target_data As Double       ' 入力データ判定用変数
    Dim categorize_count As Long    ' 判定カウント用変数
    Dim table_max As Long           ' 終端テーブル番号格納用変数
    Dim ma_count As Long            ' MAカテゴリー情報を取得
    
    ' 入力データMA操作用変数群
    Dim str_madata As Variant       ' MA回答データ格納用変数
    Dim ma_address As String        ' MA回答データアドレス格納用変数
    Dim maindata_count As Long      ' MAデータカウント変数
    
    Dim prosessing_flg As Boolean   ' ヘッダ加工判定用フラグ

    Dim pmax_number As Double       ' テーブル最大値格納用変数
    Dim pmin_number As Double       ' テーブル最小値格納用変数

    Dim str_count As Long           ' str_ctdataカウント用変数
    
    Dim identity_count As Long      ' 同一問NOカウント用変数
    Dim identity_max As Long        ' 同一問NO最大数格納用変数
    
    Dim stray_ws As Worksheet       ' 未カテゴライズデータログ出力用オブジェクト変数
    Dim ct_flg As Boolean           ' 未カテゴライズデータ判定用フラグ
    Dim stray_row As Long           ' 未カテゴライズデータログ出力アドレス用変数
    
    'Dim log_rows As Long            ' ログ出力位置格納用変数
    
    ' 初期設定
    ReDim str_address(300)
    ReDim str_qcode(300)
    ReDim str_outqcode(300)
    
    ' 画面への表示をオンにする
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "　カテゴライズ加工処理数計算中..."
    
    ' 画面への表示をオフにする
    'Application.ScreenUpdating = False
    
    
    ' 終端列番号の取得 20170502 START_ROW を6に変更
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' 未カテゴライズ情報格納用シート追加
    Set stray_ws = error_tb.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    stray_row = 2
    
    stray_ws.Name = "未カテゴライズリスト"
    stray_ws.Range("A1").Value = "SampleNo"
    stray_ws.Range("B1").Value = "QCODE"
    stray_ws.Range("C1").Value = "MA_CT"
    stray_ws.Range("D1").Value = "エラー内容"
    stray_ws.Range("E1").Value = "回答内容"
    stray_ws.Range("F1").Value = "修正内容"
    'stray_ws.Range("G1").Value = "結果"
    
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

    ' 処理件数を取得
    For column_count = 12 To column_end
    
        ' アスタリスクを見つけたとき数を数える
        ' ※アスタリスクと割当ＣＴが固定位置＆対象設問に記入あり＆Skipフラグ無し
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 1, column_count + 8).Value = "割当ＣＴ" And _
        ws_process.Cells(START_ROW, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value = "" Then
        
            
            identity_max = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row - 5
            
            ' 同一問ＮＯ分ループ
            For identity_count = 1 To identity_max
            
                ' QCODEを取得
                If ws_process.Cells(START_ROW + (identity_count - 1), column_count).Value = "" Then
                    target_coderow = Qcode_Match(ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value)
                Else
                    target_coderow = 0
                End If
            
                ' 内容が数値の時
                If target_coderow <> 0 Then
            
                    ' 件数を対象件数を増やす
                    process_max = process_max + 1
            
                    ' QCODE列番号を配列に格納する
                    str_qcode(process_max) = target_coderow
                    str_outqcode(process_max) = ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value
            
                    ' アスタリスクの位置を配列に格納する
                    str_address(process_max) = ws_process.Cells(START_ROW - 1, column_count).Address
                    
                    ' 処理内容を出力
                    Call print_log("カテゴライズ処理", ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value, _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value, _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 2).Value & "の回答を" & _
                    ws_process.Cells(START_ROW + (identity_count - 1), column_count + 3).Value & "へカテゴライズ出力。", ws_logs)

                End If
                
            Next identity_count
        
        End If
    
    Next column_count
    
    ' 格納した件数に応じて配列を再定義（値は保持する）
    ReDim Preserve str_address(process_max)
    
    ' カテゴライズ処理を行う（配列に取り込んだ指示数分）
    For process_count = 1 To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　カテゴライズ加工処理中(" & Format(process_count) & "/" & Format(process_max) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' テーブル情報を配列に格納
        
        str_ctpro = ws_process.Range(str_address(process_count)).Offset(1, 5).Resize(300, 4).Address
        str_ctdata = ws_process.Range(str_ctpro).Value2
        'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 5).Resize(300, 4).Value
        
        ' 書き込み位置を取得
        writing_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column + 1
        
        ' 入力データエリアにQCODEを設定
        ' wsp_indata.Cells(1, writing_column) = ws_process.Range(str_address(process_count)).Offset(-1, 3).Value
        
        ' ヘッダ情報を作成
        'prosessing_flg = Hedder_Create(wsp_indata, q_data(process_count).q_code _
        ', wsp_indata.Cells(1, writing_column).Address)
        
        ' ヘッダ情報を作成
        prosessing_flg = Hedder_Create(wsp_indata, str_outqcode(process_count), _
        wsp_indata.Cells(1, writing_column).Address)
        
        ' ＭＡ出力情報を取得
        ma_count = Val(ws_process.Range(str_address(process_count)).Offset(-1, 3))
        
        ' ＭＡ以外で出力を行う時
        If ma_count = 0 Then
            
            ' 初期値設定
            pmax_number = Val(str_ctdata(1, 4))
            pmin_number = Val(str_ctdata(1, 4))
            
            ' ヘッダー作成用
            For str_count = 1 To 300
            
                ' 最大値、最小値を取得
                If str_ctdata(str_count, 4) <> "" Then
                
                    ' 最大値
                    If pmax_number < Val(str_ctdata(str_count, 4)) Then
                        pmax_number = Val(str_ctdata(str_count, 4))
                    End If
                    
                    ' 最小値
                    If pmin_number > Val(str_ctdata(str_count, 4)) Then
                        pmin_number = Val(str_ctdata(str_count, 4))
                    End If
                
                Else
                    Exit For
                End If
            
            Next
            
            ' 最小値と最大値を書き込み
            wsp_indata.Cells(5, writing_column).Value = pmin_number
            wsp_indata.Cells(6, writing_column).Value = pmax_number
        
        End If
        
        ' 終端テーブル番号を取得
        table_max = UBound(str_ctdata)
        
        
        
        ' 入力データ全てにカテゴライズ処理を行う
        For indata_count = START_ROW_INDATA To indata_maxrow
            
            ' 元の回答区分によって処理をわける
            Select Case q_data(str_qcode(process_count)).q_format

                ' SAもしくは実数回答系列
                Case "S", "R", "H"
            
                    ' 入力値が存在する場合処理を行う
                    If wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value <> "" Then
                        
                        ' 未カテゴライズリスト出力用フラグ（未カテゴライズ時 False ）
                        ct_flg = False
                        
                        ' 入力データを取得
                        target_data = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
            
                        ' テーブル分の処理をまとめて行う
                        For categorize_count = 1 To table_max
                    
                            ' カテゴライズ指示が無い場合
                            If str_ctdata(categorize_count, 1) = "" Then
                                Exit For
                            End If
                    
                            ' アスタリスク箇所に情報が記入されていない時処理を行う
                            'If str_ctdata(categorize_count, 1) = "" Then
                            
                            ' 範囲指定がある場合
                            If str_ctdata(categorize_count, 3) <> "" Then
                            
                                ' テーブルの範囲内の記入値の場合
                                If target_data >= str_ctdata(categorize_count, 1) And target_data <= _
                                str_ctdata(categorize_count, 3) Then
                            
                                    ' MA出力指示を判定(SA出力)
                                    If ma_count = 0 Then
                                        
                                        ' 割当CTを出力先に設定
                                        wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                        ct_flg = True
                                        Exit For
                                        
                                    ' MA出力指示を判定 (MA出力)
                                    Else
                                        
                                        ' 割当CTを出力先に設定
                                        wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                        ct_flg = True
                                        Exit For
                            
                                    End If
                                
                                ' ※非該当(未カテゴライズ）
                                Else
                                
                                End If
                            
                            ' 始点指示のみの場合
                            Else
                                ' 次の始点がある時
                                If str_ctdata(categorize_count + 1, 1) <> "" Then
                            
                                    ' テーブルの範囲内の記入値の場合
                                    If target_data >= str_ctdata(categorize_count, 1) And target_data < _
                                    str_ctdata(categorize_count + 1, 1) Then
                            
                                        ' MA出力指示を判定(SA出力)
                                        If ma_count = 0 Then
                                        
                                            ' 割当CTを出力先に設定
                                            wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                            ct_flg = True
                                            Exit For
                                        
                                        ' MA出力指示を判定 (MA出力)
                                        Else
                                        
                                            ' 割当CTを出力先に設定
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                            ct_flg = True
                                            Exit For
                            
                                        End If
                                    
                                    ' 未カテゴライズ
                                    Else
                                        
                                    End If
                                    
                                ' 始点がない時
                                Else
                                
                                    ' テーブルの範囲内の記入値の場合
                                    If target_data >= str_ctdata(categorize_count, 1) Then
                            
                                        ' MA出力指示を判定(SA出力)
                                        If ma_count = 0 Then
                                        
                                            ' 割当CTを出力先に設定
                                            wsp_indata.Cells(indata_count, writing_column) = str_ctdata(categorize_count, 4)
                                            ct_flg = True
                                            Exit For
                                        
                                        ' MA出力指示を判定 (MA出力)
                                        Else
                                        
                                            ' 割当CTを出力先に設定
                                            wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                            ct_flg = True
                                            Exit For
                            
                                        End If
                                    
                                    ' 未カテゴライズ
                                    Else
                                        
                                    
                                    End If
                                
                                End If
                            
                            End If
                            
                            'End If
            
                        Next categorize_count
                        
                        ' カテゴライズを行わなかった場合、ログを出力
                        If ct_flg = False Then

                            ' SampleNo
                            stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
                            ' QCODE
                            stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
                            ' MA_CT
                            'stray_ws.Cells(stray_row, 3).Value = 1
                            ' エラー内容
                            stray_ws.Cells(stray_row, 4).Value = "「Table番号　" & Format(ws_process.Range(str_address(process_count)).Offset(-3, 3), "000") & "」　未カテゴライズ"
                            ' 回答内容
                            stray_ws.Cells(stray_row, 5).Value = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Value
               
                            stray_row = stray_row + 1
                            
                        End If
                    
                    ' 入力値が存在しない場合
                    Else
                    
                    End If
                
                ' MA
                Case "M"
                    
                    ' 入力データの先頭アドレスを取得
                    ma_address = wsp_indata.Cells(indata_count, q_data(str_qcode(process_count)).data_column).Address
                    
                    ' MA回答範囲を配列として取得
                    str_madata = wsp_indata.Range(ma_address).Resize(, (q_data(str_qcode(process_count)).ct_count))
                    'str_ctdata = ws_process.Range(str_address(process_count)).Offset(1, 0).Resize(300, 6)
                                        
                    ' 回答が存在している場合
                    If WorksheetFunction.Sum(str_madata) <> 0 Then
                    
                        ' MA情報をカテゴライズ
                        For maindata_count = 1 To q_data(str_qcode(process_count)).ct_count
                        
                            ' 入力データが存在している場合
                            If str_madata(1, maindata_count) > 0 Then
            
                                ' テーブル分の処理をまとめて行う
                                For categorize_count = 1 To table_max
                    
                                    ' カテゴライズ指示が無い場合
                                    If str_ctdata(categorize_count, 1) = "" Then
                                        Exit For
                                    End If
                    
                                    ' アスタリスク箇所に情報が記入されていない時処理を行う
                                    'If str_ctdata(categorize_count, 1) = "" Then
                        
                                        ' テーブルの範囲内の記入値の場合
                                        If maindata_count >= str_ctdata(categorize_count, 1) And maindata_count <= _
                                        str_ctdata(categorize_count, 3) Then
                            
                                        ' 20180629 出力先の形態に合わせて修正に変更
                                        ' カテゴライズのテーブルを逆に用意することでＭＡのシングル化に使えるため
                                            
                                            ' MAでの出力
                                            If ma_count <> 0 Then
                                            
                                                ' 割当CTを出力先に設定
                                                wsp_indata.Cells(indata_count, writing_column + Val(str_ctdata(categorize_count, 4) - 1)) = 1
                                                ct_flg = True
                                                Exit For
                                            
                                            ' MA以外での出力
                                            Else
                                            
                                                ' 割当CTを出力先に設定
                                                wsp_indata.Cells(indata_count, writing_column) = maindata_count
                                                ct_flg = True
                                                Exit For
                                            
                                            End If
                                        
                                        ' ※非該当(未カテゴライズ）テーブル範囲外の記入値の場合
                                        Else
                                        
                                            ' SampleNo
                                            stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
                                            ' QCODE
                                            stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
                                            ' MA_CT
                                            stray_ws.Cells(stray_row, 3).Value = maindata_count
                                            ' エラー内容
                                            stray_ws.Cells(stray_row, 4).Value = "「Table番号　" & Format(ws_process.Range(str_address(process_count)).Offset(-3, 3), "000") & "」　未カテゴライズ"
                                            ' 回答内容
                                            stray_ws.Cells(stray_row, 5).Value = wsp_indata.Cells(indata_count, (q_data(str_qcode(process_count)).data_column + maindata_count - 1)).Value
               
                                            stray_row = stray_row + 1
                                        
                                        End If
                                    
                                    'End If
                                
                                Next categorize_count
                            
                            ' 入力値が存在しない場合
                            Else
                            
                            End If
                        
                        Next maindata_count
                    
                    End If
   
                Case Else
                    
            End Select
            
            ' カテゴライズできなかったレコード情報を出力
            'If ct_flg = False Then
            '
            '    ' SampleNo
            '    stray_ws.Cells(stray_row, 1) = wsp_indata.Cells(indata_count, 1).Value
            '    ' QCODE
            '    stray_ws.Cells(stray_row, 2) = q_data(str_qcode(process_count)).q_code
            '    ' MA_CT
            '    'stray_ws.Cells(stray_row, 3).Value = 1
            '    ' エラー内容
            '    stray_ws.Cells(stray_row, 4).Value = "未カテゴライズ"
            '    ' 回答内容
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
' 作　成　者 村山 誠                                                           作成日  2017.04.26  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.04.26　'
' リミットマルチ加工処理用プロシージャ                                                             '
' 引数１ WorkSheet型 リミットマルチ加工処理指示シート                                              '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Limitmulti_1(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
'    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
'    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
'    Dim qcode2_dataflg As Boolean   ' 逆セット対象設問(親)判定フラグ
    Dim work_maxcol As Long         ' 指示回答終端位置

    Dim process_case As Long        ' 処理内容格納用変数
    Dim start_address As String     ' ＬＭ先頭位置アドレス格納用変数
    Dim category_count As Long      ' 回答数格納用変数
    
    Dim work_count As Long          ' 回答数カウント用変数
    
    Dim target_address As String    ' 処理レコード先頭アドレス格納用変数
    Dim ct_count As Long            ' 処理位置格納用変数
    Dim lighting_count As Long      ' 書込数カウント用変数
    
    Dim input_type As Long          ' リミットの無回答に対する入力形態判定用変数
                                    ' [1] 0 input [2] "" input
    
    ' 最大要素数を取得
    process_max = UBound(q_data, 1)
    
    ' 最大加工回数分処理を行う
    For process_count = 1 To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text _
        '& "　リミットマルチ加工処理中(" & Format(process_count - START_SUTATUSBER) & "/" & _
        'Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
        
        ' フォーマットがLMの時
        If Mid(q_data(process_count).q_format, 1, 1) = "L" Then
            
            ' 入力データ全てを判定する
            For indata_count = START_ROW_INDATA To indata_maxrow
        
                ' カウントした回答数を初期化
                lighting_count = 0
        
                ' 作業エリアの先頭アドレスを取得
                target_address = ws_process.Cells(indata_count, q_data(process_count).data_column).Address
        
                ' 対象範囲内の回答数を取得、ループカウントよりも多い要素数の設問のみに処理を行う
                If WorksheetFunction.CountIf(wsp_indata.Range(target_address).Resize(, q_data(process_count).ct_count), ">0") > _
                q_data(process_count).ct_loop Then
                
                    ' 記入内容により判定
                    Select Case q_data(process_count).q_format
                    
                        ' データをクリアする
                        Case "LC"
                        
                            ' データを初期化
                            wsp_indata.Range(target_address).Resize(, q_data(process_count).ct_count).ClearContents
        
                        ' 強番優先でカテゴリーを残す
                        Case "LA"
                        
                            ' 若番号を優先で処理を行う
                            For ct_count = q_data(process_count).ct_count To 1 Step -1
                        
                                ' ループカウント数までの回答数をカウント
                                If lighting_count < q_data(process_count).ct_loop Then
                        
                                    ' 回答内容を確認する
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' 回答数を増加させる
                                        lighting_count = lighting_count + 1
                        
                                    End If
                                    
                                ' ループカウント以降のセル
                                Else
                        
                                    ' データを初期化
                                    wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                        
                                End If
                        
                            Next ct_count
        
                        ' 若番優先でカテゴリーを残す
                        Case "L", "LM"
                            
                            ' 若番号を優先で処理を行う
                            For ct_count = 1 To q_data(process_count).ct_count
                        
                                ' ループカウント数までの回答数をカウント
                                If lighting_count < q_data(process_count).ct_loop Then
                        
                                    ' 回答内容を確認する
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' 回答数を増加させる
                                        lighting_count = lighting_count + 1
                        
                                    End If
                                    
                                ' ループカウント以降のセル
                                Else
                        
                                    ' データを初期化
                                    wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                        
                                End If
                        
                            Next ct_count
                        
                        ' 指示がない場合は何も行わない
                        Case Else
        
                    End Select
                
                End If
        
            Next indata_count
        
        End If

    Next process_count
    
End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.05.02  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.02　'
' 加工で挿入した列のヘッダ情報を作成する関数                                                       '
' 引数１ WorkSheet型 入力データシート                                                              '
' 引数２ String型 input_code                                                                       '
' 引数３ String型 hedder_address                                                                   '
' 戻り値 boolean型 処理の有無                                                                      '
'--------------------------------------------------------------------------------------------------'
Private Function Hedder_Create(ByVal wsp_indata As Worksheet, ByVal input_code As String, _
ByVal hedder_address As String) As Boolean
        
        Dim color_column As Long    ' 着色用カラム格納用変数
        Dim qcode_row As Long       ' QCODE列番号格納用変数
        Dim ma_ct As Long           ' MAカテゴリー数格納用変数
        Dim loop_count As Long      ' 処理列カウント用変数
        
        ' QCODEを検索
        qcode_row = Qcode_Match(input_code)
        
        ' MAかLMの時
        If Mid(q_data(qcode_row).q_format, 1, 1) = "M" Or Mid(q_data(qcode_row).q_format, 1, 1) = "L" Then
            ' CT数を格納
            ma_ct = q_data(qcode_row).ct_count
        Else
            ' MAとLM以外は1を設定
            ma_ct = 1
        End If
        
    ' 指定列分処理を行う
    For loop_count = 1 To ma_ct
        
        ' コメントを入力データヘッダに代入、下線を引く
        wsp_indata.Range(hedder_address).Value = input_code
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        
        ' ヘッダアドレスを変更
        hedder_address = wsp_indata.Range(hedder_address).Offset(1).Resize(3).Address
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
        
        ' 全体に罫線を引く
        hedder_address = wsp_indata.Range(hedder_address).Offset(3).Resize(2).Address
        wsp_indata.Range(hedder_address).Borders.LineStyle = xlContinuous
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).LineStyle = xlDash
        wsp_indata.Range(hedder_address).Borders(xlInsideHorizontal).Weight = xlHairline
        
        ' 範囲を変更
        hedder_address = wsp_indata.Range(hedder_address).Offset(-4).Resize(6).Address
        
        ' フォーマットによって処理を行う
        Select Case Mid(q_data(qcode_row).q_format, 1, 1)
            Case "S"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "SA"
            Case "M"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(1).Resize(1) = loop_count
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "MA"
            Case "L"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(1).Resize(1) = loop_count
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "LM"
            Case "R"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "RA"
            Case "H"
                wsp_indata.Range(hedder_address).Interior.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Interior.Color
                wsp_indata.Range(hedder_address).Font.Color = _
                ThisWorkbook.Worksheets("設定画面").Cells(qcode_row, 9).DisplayFormat.Font.Color
                wsp_indata.Range(hedder_address).Offset(3).Resize(1) = "HC"
           ' 暫定着色のため、指定する場合は仕様をきめて着色すること
           Case Else
                wsp_indata.Range(hedder_address).Interior.Color = RGB(255, 192, 0)
        End Select
        
        ' ＣＴ番号とフォーマットをセンタリング
        wsp_indata.Range(hedder_address).Offset(1).Resize(1).HorizontalAlignment = xlCenter
        wsp_indata.Range(hedder_address).Offset(3).Resize(1).HorizontalAlignment = xlCenter
        
        ' 処理位置を1行ずらす
        hedder_address = wsp_indata.Range(hedder_address).Offset(0, 1).Resize(1).Address
         
    Next loop_count

End Function

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.05.09  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.15　'
' QCODEの入力データが指示通りに記入されているかを判定するプロシージャ                              '
' 引数１ WorkSheet型 wsp_indata 入力データシート                                                   '
' 引数２ Variant型   and_data 加工指示行格納用配列変数(記入情報全て)                               '
' 引数３ Variant型   and_column 加工指示行番号格納用配列変数(Long型)                               '
' 引数４ Long型      indata_maxrow 入力データ最終列番号格納用変数                                  '
' 引数５ Long型      qcode2_row 出力先列番号格納用変数                                             '
' 戻り値 boolean型   salectflg_decision 処理の有無                                                 '
'--------------------------------------------------------------------------------------------------'
Private Sub salectflg_decision(ByVal wsp_indata As Worksheet, ByVal and_data As Variant, _
ByVal and_column As Variant, ByVal indata_maxrow As Long, ByVal qcode2_row As Long)

    Dim pindata_count As Long       ' 入力データカウント用変数
    Dim loop_count As Long          ' ループカウント用変数
    
    Dim decision_flg1 As Boolean    ' 処理判定用フラグ１
    Dim decision_flg2 As Boolean    ' 処理判定用フラグ２
    Dim decision_flg3 As Boolean    ' 処理判定用フラグ３
    
    Dim decision_type As Long       ' AND、OR条件格納用変数
    Dim qcode1_row As Long          ' 入力データ位置格納用変数
    Dim min_number As Double        ' セレクト条件最小値格納用変数
    Dim max_number As Double        ' セレクト条件最大値格納用変数
    
    Dim ma_address As String        ' 複数回答アドレス格納用変数
    
    ' 全ての入力データに処理を行う
    For pindata_count = START_ROW_INDATA To indata_maxrow
    
        'MsgBox wsp_indata.Cells(pindata_count, 1).Value
    
        ' 複数条件判定用フラグを初期化
        decision_flg1 = False
        decision_flg2 = True
        decision_flg3 = False
        
        ' 配列に格納した加工指示分の処理を行う
        For loop_count = 1 To UBound(and_column)
                
            ' 複数回答の判定条件を取得（次の条件）
            If UBound(and_data, 2) > 11 + ((loop_count - 1) * 6) + 1 And decision_flg2 = True Then
            ' 基礎情報11行、追加情報毎に6行、次の条件はその1行後

                ' OR
                If and_data(1, and_column(loop_count + 1)) = "or (もしくは)" Then
                    decision_type = 1
                ' AND
                ElseIf and_data(1, and_column(loop_count + 1)) = "and (かつ)" Then
                    decision_type = 2
                ' その他（現状次の条件が無いときのみ）
                Else
                    decision_type = 3
                End If
            
            ' 次の条件が無い時
            Else
                decision_type = 3
            End If
            
            ' QCODEの列番号を取得
            qcode1_row = Qcode_Match(and_data(1, and_column(loop_count) + 1))
            min_number = and_data(1, and_column(loop_count) + 2)
            max_number = and_data(1, and_column(loop_count) + 4)
            
            ' 指定フォーマットに合わせて処理を行う
            Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                ' 単一回答の場合
                Case "S", "R", "H"
                    
                    ' 指定のセル情報が指定の範囲の記入である時
                    If wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column) >= min_number And _
                    wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column) <= max_number Then
                    
                        ' セレクトフラグを有効
                        decision_flg1 = True    ' セレクト有効フラグ
                        
                        ' 次の接続詞に合わせてフラグを変更する、継続処理フラグを有効
                        Select Case decision_type
                        
                            ' OR条件
                            Case 1
                                decision_flg2 = True    ' 継続処理フラグ
                                decision_flg3 = True    ' OR条件開始フラグ
                            ' AND条件
                            Case 2
                                decision_flg2 = True    ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                            ' 無回答
                            Case Else
                                decision_flg2 = False   ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                        End Select
                    
                    ' 範囲外であった場合
                    Else
                    
                        ' 手前の条件がOR条件ではなく、かつ次の条件もORでは無い時
                        If decision_flg3 = False And decision_type <> 1 Then
                    
                            ' セレクト条件を満たせないためフラグを全てFALSEにし終了
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = False   ' 継続処理フラグ
                            decision_flg3 = False   ' OR条件開始フラグ
                            Exit For
                        
                        ' 手前の条件がOR以外で、かつ次の条件がORの時
                        ElseIf decision_flg3 = False And decision_type = 1 Then
                        
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = True    ' 継続処理フラグ
                            decision_flg3 = True    ' OR条件開始フラグ
                        
                        ' 手前の条件がOR条件
                        ElseIf decision_flg3 = True Then
                        
                            ' セレクト有効回答フラグがオンの時（OR条件の条件を満たしている場合
                            If decision_flg1 = True Then
                            
                                ' 次の条件がORの時
                                If decision_type = 1 Then
                            
                                    decision_flg2 = True   ' 継続処理フラグ
                                    decision_flg3 = True   ' OR条件開始フラグ
                            
                                ' 次の条件がANDの時
                                ElseIf decision_type = 2 Then
                            
                                    decision_flg2 = True   ' 継続処理フラグ
                                    decision_flg3 = False  ' OR条件開始フラグ
                                
                                ' それ以外の時
                                Else
                                
                                    decision_flg2 = False  ' 継続処理フラグ
                                    decision_flg3 = False  ' OR条件開始フラグ
                                    Exit For
                                
                                End If
                            
                            ' 次の条件がORの時
                            ElseIf decision_type = 1 Then
                                
                                decision_flg2 = True   ' 継続処理フラグ
                                decision_flg3 = True   ' OR条件開始フラグ
                                
                            Else
                                
                                decision_flg1 = False   ' セレクト有効フラグ
                                decision_flg2 = False   ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                                Exit For
                                
                            End If
                            
                        ' 次の条件がOR条件の時
                        ElseIf decision_type = 1 Then
                            
                            decision_flg2 = True   ' 継続処理フラグ
                            decision_flg3 = True   ' OR条件開始フラグ
                                
                        ' それ以外の時
                        Else
                                
                            ' セレクト条件を満たせないためフラグを全てFALSEにし終了
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = False   ' 継続処理フラグ
                            decision_flg3 = False   ' OR条件開始フラグ
                                
                        End If
                        
                    End If
                
                ' 複数回答の場合
                Case "M", "L"
                
                    ' 処理位置の先頭アドレスを格納
                    ma_address = wsp_indata.Cells(pindata_count, q_data(qcode1_row).data_column).Address
                    
                    ' 0カテゴリー時のみ参照位置を調整する
                    'If q_data(qcode1_row).ct_0flg = True Then
                    '
                    '    min_number = min_number + 1
                    '    max_number = max_number + 1
                    '
                    'End If
                    
                    ' 指定のセル情報が指定の範囲の記入である時
                    If WorksheetFunction.Sum(wsp_indata.Range(ma_address).Offset(0, min_number - 1) _
                    .Resize(, max_number - min_number + 1)) <> 0 Then
                    
                        ' セレクトフラグを有効
                        decision_flg1 = True    ' セレクト有効フラグ
                        
                        ' 次の接続詞に合わせてフラグを変更する、継続処理フラグを有効
                        Select Case decision_type
                        
                            ' OR条件
                            Case 1
                                decision_flg2 = True    ' 継続処理フラグ
                                decision_flg3 = True    ' OR条件開始フラグ
                            ' AND条件
                            Case 2
                                decision_flg2 = True    ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                            ' 無回答
                            Case Else
                                decision_flg2 = False   ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                        End Select
                    
                    ' 範囲外であった場合
                    Else
                    
                        ' 手前の条件がOR条件ではなく、かつ次の条件もORでは無い時
                        If decision_flg3 = False And decision_type <> 1 Then
                    
                            ' セレクト条件を満たせないためフラグを全てFALSEにし終了
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = False   ' 継続処理フラグ
                            decision_flg3 = False   ' OR条件開始フラグ
                            Exit For
                        
                        ' 手前の条件がOR以外で、かつ次の条件がORの時
                        ElseIf decision_flg3 = False And decision_type = 1 Then
                        
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = True    ' 継続処理フラグ
                            decision_flg3 = True    ' OR条件開始フラグ
                        
                        ' 手前の条件がOR条件
                        ElseIf decision_flg3 = True Then
                        
                            ' セレクト有効回答フラグがオンの時（OR条件の条件を満たしている場合
                            If decision_flg1 = True Then
                            
                                ' 次の条件がORの時
                                If decision_type = 1 Then
                            
                                    decision_flg2 = True   ' 継続処理フラグ
                                    decision_flg3 = True   ' OR条件開始フラグ
                            
                                ' 次の条件がANDの時
                                ElseIf decision_type = 2 Then
                            
                                    decision_flg2 = True   ' 継続処理フラグ
                                    decision_flg3 = False  ' OR条件開始フラグ
                                
                                ' それ以外の時
                                Else
                                
                                    decision_flg2 = False  ' 継続処理フラグ
                                    decision_flg3 = False  ' OR条件開始フラグ
                                    Exit For
                                
                                End If
                            
                            ' 次の条件がORの時
                            ElseIf decision_type = 1 Then
                                
                                decision_flg2 = True   ' 継続処理フラグ
                                decision_flg3 = True   ' OR条件開始フラグ
                                
                            Else
                                
                                decision_flg1 = False   ' セレクト有効フラグ
                                decision_flg2 = False   ' 継続処理フラグ
                                decision_flg3 = False   ' OR条件開始フラグ
                                Exit For
                                
                            End If
                            
                        ' 次の条件がOR条件の時
                        ElseIf decision_type = 1 Then
                            
                            decision_flg2 = True   ' 継続処理フラグ
                            decision_flg3 = True   ' OR条件開始フラグ
                                
                        ' それ以外の時
                        Else
                                
                            ' セレクト条件を満たせないためフラグを全てFALSEにし終了
                            decision_flg1 = False   ' セレクト有効フラグ
                            decision_flg2 = False   ' 継続処理フラグ
                            decision_flg3 = False   ' OR条件開始フラグ
                                
                        End If
                        
                    End If
        
                Case Else
        
            End Select
            
            ' 継続フラグが無効の場合
            If decision_flg2 = False Then
                Exit For
            End If
            
        Next loop_count
        
        ' セレクトフラグが有効の場合
        If decision_flg1 = True Then
        
            ' セレクトフラグを有効にする
            wsp_indata.Cells(pindata_count, q_data(qcode2_row).data_column).Value = 1
        
        End If
        
    Next pindata_count

    'Debug.Print and_data(1, and_column(1))

End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.05.15  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.05.19　'
' 指定された回答がある場合データをクリアする                                                       '
' 引数１ WorkSheet型 逆セット処理指示シート                                                        '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub data_clear1(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' 加工回数カウント用変数
    Dim process_max As Long             ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long              ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long              ' 逆セット対象設問(親) ROW格納用変数
    Dim input_word As String            ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean          ' 処理判定用フラグ
    
    Dim clear_min As Long               ' 削除値格納用変数（最小値）
    Dim clear_max As Long               ' 削除値格納用変数（最大値）
    
    Dim indata_count As Long            ' 入力データ処理位置格納用変数

    Dim ma_count As Long                ' MA回答内容確認用カウント変数
    
    ' 20200331 追加
    Dim amplification_count As Long     ' 増幅数カウント用変数
    
    Dim qcode_count As Long             ' 設定画面情報カウント用変数
    Dim qcode_max As Long               ' 設定画面情報最大数格納用変数
    Dim select_qcode(2, 3) As Long      ' セレクト条件QCODE格納配列
    
    Dim processing_count As Long        ' セレクト条件数カウント用変数
    Dim processing_data As Long         ' セレクト条件数格納用変数
    Dim processing_flg As Boolean       ' セレクト条件判定用変数
    Dim processing_address As String    ' セレクト条件アドレス格納用変数
    
    Dim target_data As String           ' データクリアアドレス格納用変数
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　データクリア①処理中(" & Format(process_count - START_SUTATUSBER) & _
        '"/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' スキップフラグが有効でない時、QCODEが記入されている時
        If ws_process.Cells(process_count, SKIP_FLG) = "" And _
        ws_process.Cells(process_count, QCODE1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA1) <> "" And _
        ws_process.Cells(process_count, QCODE1_DATA3) <> "" Then
        
            ' 参照設問列番号を取得
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE2).Value)
            clear_min = ws_process.Cells(process_count, QCODE1_DATA1).Value
            clear_max = ws_process.Cells(process_count, QCODE1_DATA3).Value
            
            ' 入力データ全てに処理を行う
            For indata_count = START_ROW_INDATA To indata_maxrow
            
                ' 処理判定用フラグ初期化
                process_flg = False
                
                ' 参照エリアのアドレスを取得
                target_data = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
            
                ' 参照設問のフォーマットにより処理を変更する
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' SA
                    Case "S", "R", "H"
                    
                        ' 指定範囲の記入がある場合
                        If wsp_indata.Range(target_data) >= clear_min And _
                        wsp_indata.Range(target_data) <= clear_max Then
                
                            ' 処理判定フラグを有効にする
                            process_flg = True
                
                        End If
                    
                    ' MA LM
                    Case "M", "L"
                        
                        ' ct_0flgがONの時は座標を1つ変更する
                        'If q_data(qcode1_row).ct_0flg = True Then
                        '    clear_min = clear_min + 1
                        '    clear_max = clear_max + 1
                        'End If
                        
                        ' 指定範囲の記入がある場合
                        If Application.WorksheetFunction.Sum(wsp_indata.Range(target_data). _
                        Offset(, clear_min - 1).Resize(, clear_max)) <> 0 Then
                            
                            ' 処理判定用フラグを有効にする
                            process_flg = True
                        
                        End If
                        
                    Case Else
            
                End Select
            
                ' 処理判定用フラグが有効の場合
                If process_flg = True Then
                
                    ' データクリア設問のフィーマットにより処理をわける
                    Select Case Mid(q_data(qcode2_row).q_format, 1, 1)
                        
                        ' SA
                        Case "S", "R", "H", "F", "O"
                        
                            ' データクリア設問をクリア
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
' 作　成　者 村山 誠                                                           作成日  2017.05.15  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.06.26　'
' 設置画面のセレクト条件に合わせて入力データをクリア処理                                           '
' 引数１ WorkSheet型 逆セット処理指示シート                                                        '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub data_clear2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' 加工回数カウント用変数
    Dim process_max As Long             ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long              ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long              ' 逆セット対象設問(親) ROW格納用変数
    Dim input_word As String            ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean          ' 処理判定用フラグ
    
    Dim indata_count As Long            ' 入力データ処理位置格納用変数

    Dim ma_count As Long                ' MA回答内容確認用カウント変数
    
    Dim qcode_count As Long             ' 設定画面情報カウント用変数
    Dim qcode_max As Long               ' 設定画面情報最大数格納用変数
    Dim select_qcode(2, 3) As Long      ' セレクト条件QCODE格納配列
    
    Dim processing_count As Long        ' セレクト条件数カウント用変数
    Dim processing_data As Long         ' セレクト条件数格納用変数
    Dim processing_flg As Boolean       ' セレクト条件判定用変数
    Dim processing_address As String    ' セレクト条件アドレス格納用変数
    
    Dim target_data As String           ' データクリアアドレス格納用変数
    
    ' QCODE情報を全て確認する
    For qcode_count = 1 To UBound(q_data, 1)
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text _
        '& "　データクリア②処理中(" & Format(qcode_count) & "/" & Format(UBound(q_data)) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' セレクト条件①が存在する時
        If q_data(qcode_count).sel_code1 <> "" Then
    
            ' セレクト条件に合わせクリア処理を行う
            Call select_clear(Qcode_Match(q_data(qcode_count).sel_code1), q_data(qcode_count).sel_value1, qcode_count, wsp_indata, indata_maxrow)
            
            ' セレクト条件②が存在する時
            If q_data(qcode_count).sel_code2 <> "" Then
            
                ' セレクト条件に合わせクリア処理を行う
                Call select_clear(Qcode_Match(q_data(qcode_count).sel_code2), q_data(qcode_count).sel_value2, qcode_count, wsp_indata, indata_maxrow)
                
                ' セレクト条件③が存在する時
                If q_data(qcode_count).sel_code3 <> "" Then
                
                    ' セレクト条件に合わせクリア処理を行う
                    Call select_clear(Qcode_Match(q_data(qcode_count).sel_code3), q_data(qcode_count).sel_value3, qcode_count, wsp_indata, indata_maxrow)
                
                End If
            
            End If
    
        End If
            
    Next qcode_count

End Sub

'--------------------------------------------------------------------------------------------------'
' 作　成　者 村山 誠                                                           作成日  2017.05.15  '
' 最終編集者 村山 誠　　　　　　　　　　　　　　　　　　　　　　　　　　　   　編集日　2017.06.26　'
' リミットマルチの回答状況修正処理　　  　　　　　　　　　                                         '
' 引数１ WorkSheet型 逆セット処理指示シート                                                        '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Limitmulti_2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)
    
    Dim process_count As Long           ' 加工回数カウント用変数
    Dim process_max As Long             ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long              ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long              ' 逆セット対象設問(親) ROW格納用変数
    Dim input_word As String            ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean          ' 処理判定用フラグ
    
    Dim indata_count As Long            ' 入力データ処理位置格納用変数

    Dim ma_count As Long                ' MA回答内容確認用カウント変数
    
    Dim qcode_count As Long             ' 設定画面情報カウント用変数
    Dim qcode_max As Long               ' 設定画面情報最大数格納用変数
    Dim select_qcode(2, 3) As Long      ' セレクト条件QCODE格納配列
    
    Dim processing_count As Long        ' セレクト条件数カウント用変数
    Dim processing_data As Long         ' セレクト条件数格納用変数
    Dim processing_flg As Boolean       ' セレクト条件判定用変数
    Dim processing_address As String    ' セレクト条件アドレス格納用変数
    
    Dim target_data As String           ' データクリアアドレス格納用変数
    
'    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
'    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
'    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    
'    Dim indata_count As Long        ' 入力データ処理位置格納用変数
'    Dim qcode2_dataflg As Boolean   ' 逆セット対象設問(親)判定フラグ
    Dim work_maxcol As Long         ' 指示回答終端位置

    Dim process_case As Long        ' 処理内容格納用変数
    Dim start_address As String     ' ＬＭ先頭位置アドレス格納用変数
    Dim category_count As Long      ' 回答数格納用変数
    
    Dim work_count As Long          ' 回答数カウント用変数
    
    Dim target_address As String    ' 処理レコード先頭アドレス格納用変数
    Dim ct_count As Long            ' 処理位置格納用変数
    Dim lighting_count As Long      ' 書込数カウント用変数
    
    Dim search_area As Range        ' 1.0 or 1."" 判定サーチ用変数
    Dim search_address As Range     ' サーチ情報格納用変数
    Dim search_flg As Boolean       ' サーチ判定フラグ（true = 0アリ、false = 0ナシ）
    
    ' QCODE情報を全て確認する
    For qcode_count = 1 To UBound(q_data, 1)
    
        ' フォーマットがLMの時
        If Mid(q_data(qcode_count).q_format, 1, 1) = "L" Then
        
            ' 指定のリミットマルチが1.0 or 1.""かを判定
            Set search_area = Range(ws_process.Cells(START_ROW_INDATA, q_data(qcode_count).data_column).Address, _
                ws_process.Cells(indata_maxrow, q_data(qcode_count).data_column + q_data(qcode_count).ct_count - 1).Address)
            Set search_address = search_area.Find(0, LookIn:=xlValues, lookat:=xlWhole)
            
            ' 判定フラグ切り替え
            If Not search_address Is Nothing Then
                search_flg = True
            Else
                search_flg = False
            End If
            
            ' 入力データ全てを判定する
            For indata_count = START_ROW_INDATA To indata_maxrow
        
                ' カウントした回答数を初期化
                lighting_count = 0
        
                ' 作業エリアの先頭アドレスを取得
                target_address = ws_process.Cells(indata_count, q_data(qcode_count).data_column).Address
        
                ' 対象範囲内の回答数を取得、ループカウントよりも多い要素数の設問のみに処理を行う
                If WorksheetFunction.CountIf(wsp_indata.Range(target_address).Resize(, q_data(qcode_count).ct_count), ">0") > _
                q_data(qcode_count).ct_loop Then
                
                    ' 記入内容により判定
                    Select Case q_data(qcode_count).q_format
                    
                        ' データをクリアする
                        Case "LC"
                        
                            ' データを初期化
                            wsp_indata.Range(target_address).Resize(, q_data(qcode_count).ct_count).ClearContents
        
                        ' 強番優先でカテゴリーを残す
                        Case "LA"
                        
                            ' 若番号を優先で処理を行う
                            For ct_count = q_data(qcode_count).ct_count To 1 Step -1
                        
                                ' ループカウント数までの回答数をカウント
                                If lighting_count < q_data(qcode_count).ct_loop Then
                        
                                    ' 回答内容を確認する
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' 回答数を増加させる
                                        lighting_count = lighting_count + 1
                                    
                                    ' 1.0形式の場合、0ウメ
                                    ElseIf search_flg = True Then
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1) = 0
                                    End If
                                    
                                ' ループカウント以降のセル
                                Else
                                    If search_flg = True Then
                                        ' データを初期化 1
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    Else
                                        ' データを初期化 ""
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                                    End If

                        
                                End If
                        
                            Next ct_count
        
                        ' 若番優先でカテゴリーを残す
                        Case "L", "LM"
                            
                            ' 若番号を優先で処理を行う
                            For ct_count = 1 To q_data(qcode_count).ct_count
                        
                                ' ループカウント数までの回答数をカウント
                                If lighting_count < q_data(qcode_count).ct_loop Then
                        
                                    ' 回答内容を確認する
                                    If Val(wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value) > 0 Then
                        
                                        ' 回答数を増加させる
                                        lighting_count = lighting_count + 1
                                    
                                    ' 1.0形式の場合、0ウメ
                                    ElseIf search_flg = True Then
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    End If
                                    
                                ' ループカウント以降のセル
                                Else
                                    '
                                    If search_flg = True Then
                                        ' データを初期化 1
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = 0
                                    
                                    Else
                                        ' データを初期化 ""
                                        wsp_indata.Range(target_address).Offset(0, ct_count - 1).Value = ""
                                    End If
                                    
                                End If
                        
                            Next ct_count
                        
                        ' 指示がない場合は何も行わない
                        Case Else
        
                    End Select
                
                End If
        
            Next indata_count
        
        End If
            
    Next qcode_count

End Sub

'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2017.05.08  '
' セレクトフラグ加工用プロシージャ                                                                 '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg2(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
    Dim qcode3_row As Long          ' エントリーエリア終端格納用変数
    
    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim work_maxcol As Long         ' 指示回答終端位置
    
    Dim str_ct1 As Variant          ' カテゴリー比較用配列（◎）
    Dim str_ct2 As Variant          ' カテゴリー比較用配列（○）
    Dim str_ct3 As Variant          ' カテゴリー比較用配列（○+◎）
    
    Dim ct_count As Long            ' 配列内容カウント用変数
    
'    Dim ct1_count As Long           ' ◎回答数格納用変数
'    Dim ct2_count As Long           ' ○回答数格納用変数
'    Dim ct3_count As Long           ' ○+◎回答数格納用変数
    
    Dim work1_flg As Boolean        ' 第一条件格納用フラグ
    Dim work2_flg As Boolean        ' 第二条件格納用フラグ
    
    Dim target_address As String    ' QCODE1アドレス格納用変数
    
    Dim processing_flg As Boolean   ' Function戻り値格納用変数
    
    Dim wb_calculation As Workbook  ' 中間計算用ブック情報格納用変数
    Dim ws_calculation As Worksheet ' 中間計算用シート情報格納用変数
    
    Dim start_num As Double         ' 開始番号格納用変数
    Dim end_num As Double           ' 終端番号格納用変数
    Dim connect_data As String      ' 接続詞格納用変数
    Dim target_range As Long        ' セレクト範囲取得用変数
    Dim answer_num As Long          ' 回答数格納用変数
    Dim match_data As Long          ' 作成QCODE行番号格納用変数
    
    Dim match_column As Long
    
' 各指示単一毎に条件を満たしているかを判定して別ブックに吐き出す----------------------------------------------
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 計算用ブックを作成
    Set wb_calculation = Workbooks.Add
    Set ws_calculation = wb_calculation.Worksheets(1)
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　セレクトフラグ加工準備中(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' スキップフラグが入力されていない時
        If Len(ws_process.Cells(process_count, SKIP_FLG).Value) = 0 Then
        
            ' 参照設問のQCODEの列番号を取得
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4).Value)
            qcode3_row = Qcode_Match("*加工後")
            
            ' 既に出力エリアが用意されている時（かつ加工後よりも後に出力エリアが設定されている時
            If q_data(qcode2_row).data_column <> 0 And _
            q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
        
            ' まだエリアが用意されていない時
            Else
        
                ' ヘッダを作成する
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' 新しく設定したエリアのq_dataにカラムとして設定する
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' 最小値、最大値を入力データのヘッダに格納
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' 計算シートに入力QCODEを代入する
            ws_calculation.Cells(START_ROW - 3, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1)
            
            ' 計算シートに出力QCODEを代入する
            ws_calculation.Cells(START_ROW - 2, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA4)
            
            ' 計算シートに出力先QCODE列番号を代入する
            ws_calculation.Cells(START_ROW - 1, process_count - QCODE_P_COLUMN).Value = qcode2_row
            
            ' 計算シートに集計条件を代入する
            ws_calculation.Cells(START_ROW, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA5)
            
            ' 開始番号と終端番号を変数に保持させる
            start_num = ws_process.Cells(process_count, QCODE1_DATA1).Value
            end_num = ws_process.Cells(process_count, QCODE1_DATA3).Value
            
            ' 入力データ全てに処理を行う
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' 対象設問のフォーマットにより判定条件を変更
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' 単一回答の場合
                    Case "S", "R", "H"
                
                        ' 回答が指定された範囲に含まれているか判定
                        If wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value >= start_num And _
                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <= end_num Then
                            
                            ' フラグを立てる
                            ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            
                        End If
                    
                    ' 複数回答の場合
                    Case "L", "M"
                    
                        ' 先頭アドレスを取得
                        target_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                        
                        ' 0CTフラグを判定（通常カテゴリー）
                        'If q_data(qcode1_row).ct_0flg = False Then
                        
                            ' 対象の範囲の回答を確認する
                            If WorksheetFunction.Sum(wsp_indata.Range(target_address) _
                            .Offset(0, start_num - 1).Resize(, end_num + 1 - start_num)) <> 0 Then
                        
                                ' フラグを立てる
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                            End If
                        
                        ' 0CTフラグを判定（０カテゴリー有り）
                        'Else
                        
                            ' 対象の範囲の回答を確認する
                        '    If WorksheetFunction.sum(wsp_indata.Range(target_address) _
                        '    .Offset(0, start_num).Resize(, (end_num + 1) - start_num)) <> 0 Then
                        
                                ' フラグを立てる
                        '        ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                        '    End If
                        
                        'End If

                    Case Else
                    
                    End Select
                    
            Next indata_count
        
        End If
    
    Next process_count
    
' 別ブックに吐き出した判定情報を整形しセレクトフラグを作成する---------------------------------------------
    
    ' 計算シートをアクティブに変更
    ws_calculation.Activate
    
    ' 最大加工回数分処理を行う
    For process_count = (START_ROW - QCODE_P_COLUMN) To (process_max - QCODE_P_COLUMN)
    
        ' 画面への表示をオンにする
        'Application.ScreenUpdating = True
    
        'Application.StatusBar = statusBar_text & _
        '"　セレクトフラグ加工処理中(" & Format(process_count - START_SUTATUSBER) & "/" & Format(process_max - START_SUTATUSBER) & ")"
    
        ' 画面への表示をオフにする
        'Application.ScreenUpdating = False
    
        ' 前後の指示と出力QCODEが異なる時
        If ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count + 1).Value And _
        ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count - 1).Value Then
        
            ' 計算シートをアクティブに変更
            ws_calculation.Activate
        
            ' 通常処理
            target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            
            ' 対象の範囲をコピーして入力データに貼り付ける
            ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
            Selection.Copy
            
            ' 入力データをアクティブに変更
            wsp_indata.Activate
            
            ' 書き込み位置のアドレスを取得
            target_address = wsp_indata.Cells(START_ROW_INDATA, q_data(ws_calculation.Cells(QCODE_P_ROW, process_count).Value).data_column).Address
            
            ' コピーしたデータを貼り付け
            wsp_indata.Range(target_address).PasteSpecial (xlPasteValues)
            
        ' 前の指示と出力QCODEが異なる時（始点）
        ElseIf ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count - 1).Value Then
        
            ' 始点のアドレス、指示形態を取得
            target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            connect_data = ws_calculation.Cells(QCODE_P_CA, process_count).Value
            target_range = 1
            
            ' ヘッダ情報をコピーする
            ws_calculation.Cells(QCODE_P_COLUMN, 1).Value = ws_calculation.Cells(QCODE_P_COLUMN, process_count).Value
            ws_calculation.Cells(QCODE_P_ROW, 1).Value = ws_calculation.Cells(QCODE_P_ROW, process_count).Value
        
        ' 後の指示と出力QCODEが異なる時（終点）
        ElseIf ws_calculation.Cells(4, process_count).Value <> ws_calculation.Cells(4, process_count + 1).Value Then
        
            ' カウンターを増加する
            target_range = target_range + 1
        
            ' 入力データ全てに処理を行う
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' 指定範囲の回答数を取得
                answer_num = WorksheetFunction.Sum(ws_calculation.Range(target_address) _
                .Offset(indata_count - START_ROW - 1, 0).Resize(, target_range))
            
                ' 条件に合わせて処理を行う
                Select Case connect_data
            
                    ' OR
                    Case "or (もしくは)"
                    
                        ' 対象範囲に一つでも回答があった
                        If answer_num <> 0 Then
                            
                            ' フラグを有効にする
                            ws_calculation.Cells(indata_count, 1).Value = 1
                        
                        Else
                        
                            ' 有効出ない場合はクリアを行う
                            ws_calculation.Cells(indata_count, 1).Value = ""
                        
                        End If
                        
                    ' AND
                    Case "and (かつ)"
                    
                        ' 対象範囲全てに回答があった
                        If answer_num = target_range Then
                            
                            ' フラグを有効にする
                            ws_calculation.Cells(indata_count, 1) = 1
                            
                        Else
                            
                            ' 有効出ない場合はクリアを行う
                            ws_calculation.Cells(indata_count, 1) = ""
                            
                        End If
            
                    ' 空欄だった時
                    Case Else
            
                End Select
            
            Next indata_count
            
            ' 計算シートをアクティブに変更
            ws_calculation.Activate
        
            ' 通常処理
            target_address = ws_calculation.Cells(START_ROW_INDATA, 1).Address
            
            ' 対象の範囲をコピーして入力データに貼り付ける
            ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
            Selection.Copy
            
            ' 入力データをアクティブに変更
            wsp_indata.Activate
            
            ' 書き込み位置のアドレスを取得
            target_address = wsp_indata.Cells(START_ROW_INDATA, q_data(ws_calculation.Cells(QCODE_P_ROW, 1).Value).data_column).Address
            
            ' コピーしたデータを貼り付け
            wsp_indata.Range(target_address).PasteSpecial (xlPasteValues)
            
            ' 作成したキーが他の条件に含まれている場合
            On Error Resume Next
            match_column = WorksheetFunction.Match(ws_calculation.Range("A4"), ws_calculation.Rows(3), 0)
            
            ' キーが一致した場合
            If match_column Then
                
                ' 先頭アドレスを取得
                target_address = ws_calculation.Cells(START_ROW_INDATA, match_column).Address
                
                ' コピーしたデータを貼り付け
                ws_calculation.Range(target_address).PasteSpecial (xlPasteValues)
            
            End If
            
            On Error GoTo 0
            
        ' 始点、終点を除く範囲の計算
        Else
        
            ' カウンターを増加する
            target_range = target_range + 1
        
        End If
    
    Next process_count
    
    ' オブジェクトを閉じる
    wb_calculation.Close SaveChanges:=False
    
End Sub



'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2017.06.26  '
' データクリア用プロシージャ                                                                 '
' 引数１ Long型      Select用QCODE列番号                                                           '
' 引数２ Long型      Select用Value内容                                                             '
' 引数３ Long型      QCODE処理番号格納用                                                           '
' 引数４ WorkSheet型 入力データシート                                                              '
' 引数５ Long型      入力データ終端番号                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub select_clear(ByVal select_row As Long, ByVal select_value As Long, ByVal qcode_count As Long, _
ByVal wsp_indata As Worksheet, ByVal indata_maxrow As Long)

    Dim indata_count As Long        ' 入力レコードカウント用変数
    Dim select_address As String    ' セレクトアドレス格納用変数（複数回答の時のみ使用）
    
    Dim clear_address As String     ' クリア範囲格納用変数

    ' クリア情報を取得
    Select Case q_data(qcode_count).q_format
    
        ' 単一回答
        Case "S", "F", "O"
        
            ' クリア範囲を取得
            clear_address = wsp_indata.Cells(6, q_data(qcode_count).data_column).Address
    
        ' 複数回答
        Case "M", "L", "LM", "LA", "LC"
            
            ' クリア範囲を取得
            clear_address = wsp_indata.Cells(6, q_data(qcode_count).data_column).Resize(, q_data(qcode_count).ct_count).Address
                
        ' イレギュラーのFormat
        Case Else

            clear_address = ""

    End Select

    ' クリア範囲を取得している場合
    If clear_address <> "" Then
        ' セレクト条件のフォーマットを判定
        Select Case q_data(select_row).q_format

            ' 単純集計
            Case "S", "F", "O"
            
                ' 0カテゴリーもそのまま処理
                For indata_count = START_ROW_INDATA To indata_maxrow
            
                    ' 回答内容が一致した時
                    If wsp_indata.Cells(indata_count, q_data(select_row).data_column).Value = select_value Then
                
                        ' 処理を行わない
                
                    ' 回答内容が一致しなかった時
                    Else
                    
                        ' 範囲をクリア
                        wsp_indata.Range(clear_address).Offset(indata_count - 6).ClearContents
                
                    End If
            
                Next indata_count
            
            ' 複数回答関連
            Case "M", "L", "LM", "LA", "LC"
        
                ' 入力データを判定
                For indata_count = START_ROW_INDATA To indata_maxrow
                    
                    ' 指定の記入値がある時
                    If Val(wsp_indata.Cells(indata_count, q_data(select_row).data_column).Offset(, select_value).Value) > 0 Then
                    ' 何も行わない
                    ' 指定の記入値がない時
                    Else
                            ' 入力値をクリア
                        wsp_indata.Range(clear_address).Offset(indata_count - 6, 0).ClearContents
                    End If
            
                Next indata_count

            ' イレギュラーのFormat
            Case Else

        End Select

    End If

End Sub




'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2020.03.30  '
' 増幅加工用プロシージャ                                                                           '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub amplification_data(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim str_address() As String     ' 処理アドレス格納用配列
    Dim str_qcode() As Long         ' 処理QCODE格納用配列
    Dim str_outqcode() As String    ' 出力QCODE格納用配列
    Dim target_coderow As Long      ' 対象QCODE列番号一時格納用変数
    
    Dim column_count As Long        ' 処理列番号格納用変数
    Dim column_end As Long          ' 列終端番号格納用変数
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim rows_count As Long          ' 入力位置格納用変数

    Dim str_ctdata As Variant       ' カテゴライズテーブル情報格納用配列変数
    
    'Dim writing_column As Long      ' 書き込み位置格納用変数
    'Dim target_data As Double       ' 入力データ判定用変数
    'Dim categorize_count As Long    ' 判定カウント用変数
    'Dim table_max As Long           ' 終端テーブル番号格納用変数
    'Dim ma_count As Long            ' MAカテゴリー情報を取得
    
    ' 入力データMA操作用変数群
    Dim str_madata As Variant       ' MA回答データ格納用変数
    Dim ma_address As String        ' MA回答データアドレス格納用変数
    Dim maindata_count As Long      ' MAデータカウント変数
    
    Dim prosessing_flg As Boolean   ' ヘッダ加工判定用フラグ

    Dim pmax_number As Double       ' テーブル最大値格納用変数
    Dim pmin_number As Double       ' テーブル最小値格納用変数

    Dim str_count As Long           ' str_ctdataカウント用変数
    
    Dim identity_count As Long      ' 同一問NOカウント用変数
    Dim identity_max As Long        ' 同一問NO最大数格納用変数
    
    Dim stray_ws As Worksheet       ' 未カテゴライズデータログ出力用オブジェクト変数
    Dim ct_flg As Boolean           ' 未カテゴライズデータ判定用フラグ
    
    Dim amp_qcode As String         ' 増幅設問名格納用変数
    
    Dim table_count As Long         ' テーブルアドレス情報カウント用変数
    Dim table_max As Long           ' 最大テーブル数格納用変数
    
    Dim target_table As Variant     ' 処理テーブル情報格納用変数
    Dim table_count_y As Long       ' 格納テーブル情報参照用変数（縦軸）
    Dim table_count_x As Long       ' 格納テーブル情報参照用変数（横軸）
    
    
    'Dim log_rows As Long            ' ログ出力位置格納用変数
    
    ' 初期設定
    ReDim str_address(300)
    ReDim str_qcode(300)
    ReDim str_outqcode(300)
    
    ' 画面への表示をオンにする
    'Application.ScreenUpdating = True
    
    'Application.StatusBar = statusBar_text & "　カテゴライズ加工処理数計算中..."
    
    ' 画面への表示をオフにする
    'Application.ScreenUpdating = False
    
    
    ' 終端列番号の取得 20170502 START_ROW を6に変更
    column_end = ws_process.Cells(START_ROW - 1, Columns.Count).End(xlToLeft).Column
    
    ' カウント数を初期化
    table_count = 1
    
    ' 未カテゴライズ情報格納用シート追加
    'Set stray_ws = error_tb.Worksheets.Add(after:=Worksheets(Worksheets.Count))
    
    'stray_ws.Name = "未カテゴライズリスト"
    'stray_ws.Range("A1").Value = "SampleNo"
    'stray_ws.Range("B1").Value = "QCODE"
    'stray_ws.Range("C1").Value = "MA_CT"
    'stray_ws.Range("D1").Value = "エラー内容"
    'stray_ws.Range("E1").Value = "回答内容"
    'stray_ws.Range("F1").Value = "修正内容"
    'stray_ws.Range("G1").Value = "結果"

    ' 増幅レコード数を取得
    For column_count = 12 To column_end
    
        ' アスタリスクを見つけたとき数を数える
        ' ※アスタリスクと割当ＣＴが固定位置＆対象設問に記入あり＆Skipフラグ無し
        If ws_process.Cells(START_ROW - 1, column_count).Value = "*" And _
        ws_process.Cells(START_ROW - 2, column_count + 3).Value <> "" And _
        ws_process.Cells(START_ROW, column_count + 2).Value <> "" And _
        ws_process.Cells(START_ROW - 2, column_count + 2).Value = "" Then
            
            ' アドレスを取得
            str_address(table_count) = ws_process.Cells(START_TABLE_DATA, column_count).Address
            table_count = table_count + 1
            
            ' 項目数をカウント
            identity_count = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row - 5
            
            ' カウントした項目数がもっとも多い場合、項目数を取得
            If identity_max < identity_count Then
                identity_max = identity_count
            End If
            
            'str_address = ws_process.Cells(Rows.Count, column_count + 2).End(xlUp).Row
            
        End If
    
    Next column_count
    
    ' 最大テーブル数を取得
    table_max = table_count
    
    ' 入力データすべてに、全テーブル分の処理を行う
    ' 　⇒　回答を増幅、テーブル数分の処理を行う
    For indata_count = START_ROW_INDATA To indata_maxrow
    
        ' 20200402 追記
        'wsp_indata.Rows (indata_count)
    
        ' １レコードをコピー
        wsp_indata.Rows(indata_count).Copy
        ' 貼り付け先の座標を取得
        rows_count = wsp_indata.Cells(Rows.Count, 1).End(xlUp).Row + 1
        ' データの増幅
        wsp_indata.Rows(rows_count & ":" & (rows_count + identity_count - 1)).PasteSpecial
        
    Next indata_count
    
    ' テーブル数分処理を行う
    For table_count = 1 To table_max
        
        ' テーブル情報をすべて格納
        target_table = ws_process.Range(str_address(table_count)).Resize(299, 12).Value
        
        ' 縦軸ループ
        For table_count_y = 1 To identity_max
        
            ' Skipフラグがたっていない場合
            If target_table(table_count_y, 1) = "" Then
        
                ' 横軸ループ
                For table_count_x = 3 To identity_max
            
                    ' 対象問Ｎｏをクリアor代入処理
                    Qcode_Match (ws_process.Cells(table_count_y, table_count_x).Value)
                    
            
                Next table_count_x
        
            End If
        
        Next table_count_y
        
        
        
    Next table_count
    
End Sub



'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2018.06.14  '
' ログ格納用プロシージャ                                                                           '
' 引数１ String型    加工内容データ                                                                '
' 引数２ String型    QCODE1データ                                                                  '
' 引数３ String型    QCODE2データ                                                                  '
' 引数４ String型    処理内容データ                                                                '
' 引数５ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub print_log(ByVal log_data1 As String, ByVal log_data2 As String, ByVal log_data3 As String, _
ByVal log_data4 As String, ByVal ws_logs As Worksheet)

    Dim row_data As Long    ' 最終行格納用変数

    ' 最終行取得
    row_data = ws_logs.Cells(Rows.Count, 1).End(xlUp).Row

    ' SEQ
    ws_logs.Cells(row_data + 1, 1) = Format(row_data)
    
    ' その他データの出力
    ws_logs.Cells(row_data + 1, 2) = log_data1
    ws_logs.Cells(row_data + 1, 3) = log_data2
    ws_logs.Cells(row_data + 1, 4) = log_data3
    ws_logs.Cells(row_data + 1, 5) = log_data4
    
End Sub

'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2018.06.29  '
' ログ格納用プロシージャ（カテゴライズ用）                                                         '
' 引数１ Long型      レコード情報                                                                  '
' 引数２ String型    QCODE1データ                                                                  '
' 引数３ Long型      MA_CT                                                                         '
' 引数４ String型    回答内容データ                                                                '
' 引数５ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub print_log2(ByVal sample_row As Long, ByVal qcode_data As String, ByVal ma_ct As Long, _
ByVal input_data As String, ByVal ws_logs As Worksheet)

    Dim row_data As Long    ' 最終行格納用変数

    ' 最終行取得
    row_data = ws_logs.Cells(Rows.Count, 1).End(xlUp).Row

    ' SEQ
    'ws_logs.Cells(row_data + 1, 1) = Format(row_data)
    
    ' その他データの出力
    'ws_logs.Cells(row_data + 1, 2) = log_data1
    'ws_logs.Cells(row_data + 1, 3) = log_data2
    'ws_logs.Cells(row_data + 1, 4) = log_data3
    'ws_logs.Cells(row_data + 1, 5) = log_data4
    
End Sub


'--------------------------------------------------------------------------------------------------'
' 作成者  村山誠                                                               作成日  2017.05.08  '
' セレクトフラグ加工用プロシージャ                                                                 '
' 引数１ WorkSheet型 カテゴライズ処理指示シート                                                    '
' 引数２ WorkSheet型 入力データシート                                                              '
' 引数３ Long型      入力データ終端番号                                                            '
' 引数４ WorkSheet型 加工ログ出力シート                                                            '
'--------------------------------------------------------------------------------------------------'
Private Sub Processing_Selectflg3(ByVal ws_process As Worksheet, ByVal wsp_indata As Worksheet, _
ByVal indata_maxrow As Long, ByVal statusBar_text As String, ByVal ws_logs As Worksheet)

    Dim process_count As Long       ' 加工回数カウント用変数
    Dim process_max As Long         ' 最大加工回数格納用変数
    
    Dim qcode1_row As Long          ' 比較設問(子) ROW格納用変数
    Dim qcode2_row As Long          ' 逆セット対象設問(親) ROW格納用変数
    Dim qcode3_row As Long          ' エントリーエリア終端格納用変数
    
    Dim input_word As String        ' 逆セット用InputWord格納用文字列変数
    Dim process_flg As Boolean      ' 処理判定用フラグ
    
    Dim indata_count As Long        ' 入力データ処理位置格納用変数
    Dim work_maxcol As Long         ' 指示回答終端位置
    
    Dim str_ct1 As Variant          ' カテゴリー比較用配列（◎）
    Dim str_ct2 As Variant          ' カテゴリー比較用配列（○）
    Dim str_ct3 As Variant          ' カテゴリー比較用配列（○+◎）
    
    Dim ct_count As Long            ' 配列内容カウント用変数
    
'    Dim ct1_count As Long           ' ◎回答数格納用変数
'    Dim ct2_count As Long           ' ○回答数格納用変数
'    Dim ct3_count As Long           ' ○+◎回答数格納用変数
    
    Dim work1_flg As Boolean        ' 第一条件格納用フラグ
    Dim work2_flg As Boolean        ' 第二条件格納用フラグ
    
    Dim target_address As String    ' QCODE1アドレス格納用変数
    
    Dim processing_flg As Boolean   ' Function戻り値格納用変数
    
    Dim wb_calculation As Workbook  ' 中間計算用ブック情報格納用変数
    Dim ws_calculation As Worksheet ' 中間計算用シート情報格納用変数
    
    Dim start_num As Double         ' 開始番号格納用変数
    Dim end_num As Double           ' 終端番号格納用変数
    Dim connect_data As String      ' 接続詞格納用変数
    Dim target_range As Long        ' セレクト範囲取得用変数
    Dim answer_num As Long          ' 回答数格納用変数
    Dim match_data As Long          ' 作成QCODE行番号格納用変数
    
    Dim work_flg As Long            ' 複数条件処理フラグ
    Dim work_address As String      ' 複数条件範囲格納用変数
    Dim match_column As Long
    
' 各指示単一毎に条件を満たしているかを判定して別ブックに吐き出す----------------------------------------------
    
    ' 加工処理数を取得
    process_max = ws_process.Cells(Rows.Count, QCODE1).End(xlUp).Row
    
    ' 計算用ブックを作成
    Set wb_calculation = Workbooks.Add
    Set ws_calculation = wb_calculation.Worksheets(1)
    
    ' 最大加工回数分処理を行う
    For process_count = START_ROW To process_max
    
        ' スキップフラグが入力されていない時
        If Len(ws_process.Cells(process_count, SKIP_FLG).Value) = 0 Then
        
            ' 参照設問のQCODEの列番号を取得
            qcode1_row = Qcode_Match(ws_process.Cells(process_count, QCODE1).Value)
            
            ' 初期データではない・QCODE2が無回答ではない
            If ws_process.Cells(process_count, QCODE1_DATA4).Value <> "" Then
                qcode2_row = Qcode_Match(ws_process.Cells(process_count, QCODE1_DATA4).Value)
            ' QCODE2が無回答の時、複数条件フラグが有効の場合は手前の出力設定をコピー
            ElseIf process_count <> START_ROW And ws_process.Cells(process_count, QCODE1_DATA4).Value = "" Then
                qcode2_row = Qcode_Match(ws_process.Cells(process_count - 1, QCODE1_DATA4).Value)
                ws_process.Cells(process_count, QCODE1_DATA4).Value = ws_process.Cells(process_count - 1, QCODE1_DATA4).Value
            
            Else
            ' 初期データもしくは出力設問番号が無回答の時
                
            End If
            
            qcode3_row = Qcode_Match("*加工後")
            
            ' 出力設問番号を判定し、ヘッダーを作成する
            If q_data(qcode2_row).data_column > q_data(qcode3_row).data_column Then
            
                ' 通常処理
            
            Else
            
                ' ヘッダを作成する
                processing_flg = Hedder_Create(wsp_indata, ws_process.Cells(process_count, QCODE1_DATA4), _
                wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Offset(, 1).Address)
                
                ' 新しく設定したエリアのq_dataにカラムとして設定する
                q_data(qcode2_row).data_column = wsp_indata.Cells(1, Columns.Count).End(xlToLeft).Column
                
                ' 最小値、最大値を入力データのヘッダに格納
                wsp_indata.Cells(5, q_data(qcode2_row).data_column) = 1
                wsp_indata.Cells(6, q_data(qcode2_row).data_column) = 1
            
            End If
            
            ' 計算シートに入力QCODEを代入する
            ws_calculation.Cells(START_ROW - 3, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1)
            
            ' 計算シートに出力QCODEを代入する
            ws_calculation.Cells(START_ROW - 2, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA4)
            
            ' 計算シートに出力先QCODE列番号を代入する
            ws_calculation.Cells(START_ROW - 1, process_count - QCODE_P_COLUMN).Value = qcode2_row
            
            ' 計算シートに集計条件を代入する
            ws_calculation.Cells(START_ROW, process_count - QCODE_P_COLUMN).Value = _
            ws_process.Cells(process_count, QCODE1_DATA5)

            ' 開始番号と終端番号を変数に保持させる（ショートカット以外）
            If ws_process.Cells(process_count, QCODE1_DATA1).Value <> "*" And _
            ws_process.Cells(process_count, QCODE1_DATA1).Value <> "_" Then
                start_num = ws_process.Cells(process_count, QCODE1_DATA1).Value
                end_num = ws_process.Cells(process_count, QCODE1_DATA3).Value
                
                ' 最小値・最大値を設定
                ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = start_num
                ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = end_num
            
            ' ショートカット使用時は初期化
            Else
                start_num = 0
                end_num = 0
                
                If ws_process.Cells(process_count, QCODE1_DATA1).Value = "*" Then
                
                    ' 最小値・最大値を設定
                    ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = "*"
                    ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = "*"
            
                ElseIf ws_process.Cells(process_count, QCODE1_DATA1).Value = "_" Then
                
                    ' 最小値・最大値を設定
                    ws_calculation.Cells(START_ROW - 5, process_count - QCODE_P_COLUMN).Value = "_"
                    ws_calculation.Cells(START_ROW - 4, process_count - QCODE_P_COLUMN).Value = "_"
                
                End If
            
            End If

            ' 入力データ全てに処理を行う
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' 対象設問のフォーマットにより判定条件を変更
                Select Case Mid(q_data(qcode1_row).q_format, 1, 1)
                
                    ' 単一回答の場合
                    Case "S", "R", "H"
                
                        ' 何かしら回答がある時
                        If ws_process.Cells(process_count, QCODE1_DATA1).Value = "*" Then
                            
                            If Len(Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value)) <> 0 Then
                            
                                ' フラグを立てる
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            End If
                        
                        ' 何も回答がない時
                        ElseIf ws_process.Cells(process_count, QCODE1_DATA1).Value = "_" Then
                            
                            If Len(Trim(wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value)) = 0 Then
                            
                                ' フラグを立てる
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            End If
                            
                        ' 指定範囲に回答がある時
                        ElseIf wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value >= start_num And _
                        wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Value <= end_num Then
                            
                            ' フラグを立てる
                            ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                            
                        End If
                    
                    ' 複数回答の場合
                    Case "L", "M"
                    
                        ' 先頭アドレスを取得
                        target_address = wsp_indata.Cells(indata_count, q_data(qcode1_row).data_column).Address
                            
                            ' 対象の範囲の回答を確認する
                            If WorksheetFunction.Sum(wsp_indata.Range(target_address) _
                            .Offset(0, start_num - 1).Resize(, end_num + 1 - start_num)) <> 0 Then
                        
                                ' フラグを立てる
                                ws_calculation.Cells(indata_count, process_count - QCODE_P_COLUMN).Value = 1
                        
                            End If

                    Case Else
                    
                    End Select
                    
            Next indata_count
        
        End If
    
    Next process_count
    
' 別ブックに吐き出した判定情報を整形しセレクトフラグを作成する---------------------------------------------
    
    ' 計算シートをアクティブに変更
    ws_calculation.Activate
    
    ' 最大加工回数分処理を行う
    For process_count = (START_ROW - QCODE_P_COLUMN) To (process_max - QCODE_P_COLUMN)
        
        ' 複数条件フラグを未使用
        If ws_calculation.Cells(6, process_count).Value = "" Then
        
            ' 開始番号・終端番号が消されていない時
            If ws_calculation.Cells(1, process_count).Value <> "" Then
        
                ' 計算シートをアクティブに変更
                ws_calculation.Activate
        
                ' 通常処理
                target_address = ws_calculation.Cells(START_ROW_INDATA, process_count).Address
            
                ' 対象の範囲をコピーして入力データに貼り付ける
                ws_calculation.Range(target_address).Resize(indata_maxrow - START_ROW).Select
                Selection.Copy
            
                ' 入力データをアクティブに変更
                wsp_indata.Activate
            
                ' 書き込み位置のアドレスを取得
                target_address = wsp_indata.Cells(START_ROW_INDATA, _
                q_data(ws_calculation.Cells(QCODE_P_ROW, process_count).Value).data_column).Address
            
                ' コピーしたデータを貼り付け(空白を無視して張り付ける)
                wsp_indata.Range(target_address).PasteSpecial xlPasteValues, SkipBlanks:=True
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then
                
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "～" & _
                    ws_calculation.Cells(2, process_count).Value & "の回答内容を出力。", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "有効回答で出力。", ws_logs)
                ElseIf ws_calculation.Cells(1, process_count).Value = "_" Then
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "無効回答で出力。", ws_logs)
                End If
                
            
            ' エラーフラグが立っている時
            Else
            
                ' エラー出力
                Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                ws_calculation.Cells(3, process_count - 1).Value & "との「" & _
                ws_calculation.Cells(START_ROW, process_count - 1).Value & _
                "」条件の指示に対し、異なる出力設問番号が設定されています。", ws_logs)
            
            End If
            
        ' 複数条件がある時、出力設定番号が同一の時（処理進行）
        ElseIf ws_calculation.Cells(4, process_count).Value = ws_calculation.Cells(4, process_count + 1).Value Then
        
            
            ' 複数条件の内容を代入（ＡＮＤ条件）
            If Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "an" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "An" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "AN" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "ａｎ" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "ＡＮ" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "Ａｎ" Then
            
                work_flg = 1
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then
                
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "～" & _
                    ws_calculation.Cells(2, process_count).Value & "の回答内容を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＡＮＤ条件で出力設定。", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "有効回答を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＡＮＤ条件で出力設定。", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "_" Then
                    
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "無効回答を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＡＮＤ条件で出力設定。", ws_logs)
                
                Else
                
                End If
                
            
            ' 複数条件の内容を代入（ＯＲ条件）
            ElseIf Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "or" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "Or" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "OR" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "ｏｒ" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "ＯＲ" Or _
            Mid(Trim(ws_calculation.Cells(6, process_count).Value), 1, 2) = "Ｏｒ" Then
            
                work_flg = 2
                
                If ws_calculation.Cells(1, process_count).Value <> "*" And _
                ws_calculation.Cells(1, process_count).Value <> "_" Then

                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    ws_calculation.Cells(1, process_count).Value & "～" & _
                    ws_calculation.Cells(2, process_count).Value & "の回答内容を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＯＲ条件で出力設定。", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                    
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "有効回答を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＯＲ条件で出力設定。", ws_logs)
                
                ElseIf ws_calculation.Cells(1, process_count).Value = "*" Then
                
                    ' 処理内容を出力
                    Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
                    ws_calculation.Cells(START_ROW - 2, process_count).Value, _
                    "無効回答を" & _
                    ws_calculation.Cells(3, process_count + 1) & "とＯＲ条件で出力設定。", ws_logs)
                
                Else
                
                End If
                
            ' 指示以外の情報が入っている場合
            Else
            
                work_flg = 0
            
            End If
            
            ' 入力データ全てに処理を行う
            For indata_count = START_ROW_INDATA To indata_maxrow
                
                ' アドレスで位置情報を取得（ワークシートファンクションを使用するため）
                work_address = ws_calculation.Cells(indata_count, process_count).Address
                
                ' ＡＮＤ条件の時
                If work_flg = 1 Then
                    
                    ' フラグの合計が2以上の時
                    If WorksheetFunction.Sum(ws_calculation.Range(work_address).Resize(, 2).Value) = 2 Then
                        ws_calculation.Range(work_address).Offset(, 1).Value = 1
                    Else
                        ws_calculation.Range(work_address).Offset(, 1).Value = ""
                    End If
                
                ' ＯＲ条件の時
                ElseIf work_flg = 2 Then
                    
                    ' フラグの合計が2以上の時
                    If WorksheetFunction.Sum(ws_calculation.Range(work_address).Resize(, 2).Value) > 0 Then
                        ws_calculation.Range(work_address).Offset(, 1).Value = 1
                    Else
                        ws_calculation.Range(work_address).Offset(, 1).Value = ""
                    End If
                
                ' 上記以外の時
                Else
                
                    ' 手前で判定をしているのでここに入ることはない予定
                    ' 可能性としてはエントリーエリアへの指示の時に通る可能性があり
                
                End If

            Next indata_count

            
        ' 複数条件の設定があるにも関わらず出力設問番号が異なる時
        Else
        
            ' エラー出力
            Call print_log("セレクトフラグ処理", ws_calculation.Cells(3, process_count), _
            ws_calculation.Cells(START_ROW - 2, process_count).Value, _
            ws_calculation.Cells(3, process_count + 1).Value & "との「" & _
            ws_calculation.Cells(START_ROW, process_count).Value & _
            "」条件の指示に対し、異なる出力設問番号が設定されています。", ws_logs)
            
            ' 処理判定用の開始番号・終端番号を消す
            ws_calculation.Cells(START_ROW - 4, process_count).Value = ""
            ws_calculation.Cells(START_ROW - 5, process_count).Value = ""
            
            ws_calculation.Cells(START_ROW - 4, process_count + 1).Value = ""
            ws_calculation.Cells(START_ROW - 5, process_count + 1).Value = ""
            
        End If
    
    Next process_count
    
    ' オブジェクトを閉じる
    wb_calculation.Close SaveChanges:=False
    
End Sub



