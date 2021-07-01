'-----------------------------------------------------------[date: 2021.07.01]
Attribute VB_Name = "Module1"
Option Explicit

Public Sub test01()
  Dim ws As Worksheet
  For Each ws In ActiveWorkbook.Worksheets
    Debug.Print ws.Name
  Next ws
End Sub
'-----------------------------------------------------------------------------


