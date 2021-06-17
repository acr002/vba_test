'-----------------------------------------------------------[date: 2021.06.17]
Attribute VB_Name = "Run"
Option Explicit

' 空のシートをusedrangeで取得するとemptyが入っています。(代入の失敗だと思われます)
' 1セルしか使用していない場合は、stringなどが入ります。
Public Sub test01()
  Dim a As Variant
  a = ActiveSheet.UsedRange.Value
  'Debug.Print a
  If Not IsEmpty(a) Then
    If IsArray(a) Then
      Debug.Print UBound(a, 1)
      Debug.Print UBound(a, 2)
    Else
      Debug.Print a
    End If
  End If
  Cells(10, 1).Resize(UBound(a, 1), UBound(a, 2)) = a
End Sub
'-----------------------------------------------------------------------------





