'-----------------------------------------------------------[date: 2021.06.17]
Attribute VB_Name = "Run"
Option Explicit

' ��̃V�[�g��usedrange�Ŏ擾�����empty�������Ă��܂��B(����̎��s���Ǝv���܂�)
' 1�Z�������g�p���Ă��Ȃ��ꍇ�́Astring�Ȃǂ�����܂��B
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





