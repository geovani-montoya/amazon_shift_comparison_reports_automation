Attribute VB_Name = "deleteConnections"
Option Explicit

Sub delConnect()
    Dim xConnect As Object
    For Each xConnect In ActiveWorkbook.Connections
    If xConnect.Name <> "ThisWorkbookDataModel" Then xConnect.Delete
    Next xConnect
End Sub


