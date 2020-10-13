Attribute VB_Name = "copy_report"
Sub copy_report()
'Copies report to new sheet with name of choice
'suggested to use date of the first day
    Dim newSheet_name As Variant
    
    Call InitializeVariables
    
    newSheet_name = InputBox("Name the new worksheet (suggested to name by the first date on the week)")


    Sheets("Report Generator").Select
    Sheets("Report Generator").Copy Before:=Sheets("Report Generator")
    Sheets("Report Generator (2)").Select

    Sheets("Report Generator (2)").Name = newSheet_name
    Sheets(newSheet_name).Select


End Sub



