'-----------------------------------------------------------[date: 2021.07.05]
Attribute VB_Name = "Module1"
Option Explicit

' シートモジュールに設置してください。
' 下記では17列目以降が選択された場合は次の行の先頭にカーソルを移動します。
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Column >= 17 Then
    Cells(Target.Row + 1, 1).Select
  End If
End Sub

