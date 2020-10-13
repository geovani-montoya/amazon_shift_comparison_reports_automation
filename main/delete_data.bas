Attribute VB_Name = "delete_data"
Sub reset_report()
Attribute reset_report.VB_ProcData.VB_Invoke_Func = " \n14"
'This clears the data to recycle the report
    Application.ScreenUpdating = Flase
    Range("B14:P20").Select
    Selection.ClearContents
    Application.ScreenUpdating = True
    Sheets("Report Generator").Range("D2").Select
    
End Sub


   

