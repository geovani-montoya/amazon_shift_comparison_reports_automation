Attribute VB_Name = "delSheets"
Option Explicit

Sub delShts()
    Dim wb As Workbook
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    Set wb = ThisWorkbook
    For Each ws In wb.Worksheets
     Select Case UCase(Left(ws.Name, 1))
     Case "p"
        ws.Delete
     Case Else
        ' do nothing
     End Select
    Next ws
     Application.DisplayAlerts = True
End Sub


