'-----------------------------------------------------------[date: 2021.07.05]
Attribute VB_Name = "Module1"
Option Explicit

' �V�[�g���W���[���ɐݒu���Ă��������B
' ���L�ł�17��ڈȍ~���I�����ꂽ�ꍇ�͎��̍s�̐擪�ɃJ�[�\�����ړ����܂��B
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Column >= 17 Then
    Cells(Target.Row + 1, 1).Select
  End If
End Sub

