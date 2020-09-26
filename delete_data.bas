Attribute VB_Name = "delete_data"
Sub restart_report()
Attribute restart_report.VB_ProcData.VB_Invoke_Func = " \n14"
'This clears the data to recycle the report

    Range("B14:P20,B34:P40,A14:A20,A34:A40").Select
    Range("B34").Activate
    Selection.ClearContents
    Range("C2").Select
End Sub


   

