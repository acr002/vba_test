'-----------------------------------------------------------[date: 2021.05.25]
Attribute VB_Name = "Module1"
Option Explicit

Public Sub test01()
  Dim dic As Variant
  Set dic = CreateObject("scripting.dictionary")
  dic("A") = "test"
  Debug.Print dic("A"), 1
  Debug.Print dic("B"), 2
  Debug.Print dic.Count
End Sub
'-----------------------------------------------------------------------------

Public Sub test02()
  Dim dic As Variant
  Set dic = CreateObject("scripting.dictionary")
  dic("A") = "test"
  If dic.exists("A") Then
    Debug.Print dic("A"), 1
  End If
  If dic.exists("B") Then
    Debug.Print dic("B"), 2
  End If
  Debug.Print dic.Count
End Sub
'-----------------------------------------------------------------------------





