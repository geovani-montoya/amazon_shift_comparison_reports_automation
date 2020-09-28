Attribute VB_Name = "copy_report"
Sub copy_report()
'Copies report to new sheet with name of choice
'suggested to use date of the first day
    Dim newSheet_name As Variant
    
    Call InitializeVariables
    
    newSheet_name = InputBox("Name the new worksheet (suggested to name by the first date on the week)")

    Sheets("Report Generator").Select
    Sheets("Report Generator").Copy After:=Sheets(Sheets.Count)
    Sheets("Report Generator (2)").Select
    
        
    Sheets("Report Generator (2)").Shapes.Range(Array("Rectangle 3")).Select
    Sheets("Report Generator (2)").Shapes.Range(Array("Rectangle 3", "Rectangle 1")).Select
    Sheets("Report Generator (2)").Shapes.Range(Array("Rectangle 3", "Rectangle 1", "Rectangle 8") _
        ).Select
    Sheets("Report Generator (2)").Shapes.Range(Array("Rectangle 3", "Rectangle 1", "Rectangle 8", _
        "Rectangle 7")).Select
    Selection.Delete
    Range("C2,Q3,Q3,F4,B4,B5:D5").Select
    Range("B5").Activate
    Selection.ClearContents
    'Sheets("Report Generator (2)").SmallScroll Down:=-3
    Range("A1").Select
    

    
    
    Sheets("Report Generator (2)").Name = newSheet_name

End Sub



