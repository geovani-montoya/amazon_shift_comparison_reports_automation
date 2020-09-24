Attribute VB_Name = "variables"
Public wbMain As Workbook
Public shtMain As Worksheet

Public Sub InitializeVariables()
    Dim myValue As Variant
    'Dim Sheet As Worksheet
    
    'myValue = InputBox("Give new worksheet title (data format is suggested)")

    Set wbMain = ActiveWorkbook
    'Set Sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    'ActiveSheet.Name = myValue
    
    With wbMain
        Set shtMain = .Sheets("demo")
        'Set shtMain = .Sheets(myValue)
    End With
   
    
End Sub

'Use this function to work with existing worksheets or new worksheets without errors

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook

    Dim Sheet As Object
    For Each Sheet In InWorkbook.Sheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
    sheetExists = False
End Function


