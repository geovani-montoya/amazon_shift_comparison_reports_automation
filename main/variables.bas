Attribute VB_Name = "variables"
Public wbMain As Workbook
Public shtMain As Worksheet

Public Sub InitializeVariables()
    Dim myValue As Variant
    
    myValue = InputBox("Give new worksheet title (data format is suggested)")

    Set wbMain = ActiveWorkbook
    
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
    ActiveSheet.Name = myValue
    
    With wbMain
        'Set shtMain = .Sheets("demo")
        Set shtMain = .Sheets(myValue)
    End With
   
    
End Sub


