Attribute VB_Name = "Main_Process"
Sub Main_Process()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False

    Dim folder_path As String, file_extension As String, input_file As String
    Dim wbPPR As Workbook, wbPID As Workbook, wbLPIstow As Workbook, wbMTpick As Workbook
    Dim wbFRR As Workbook, wbLPIpick As Workbook, wbUR As Workbook ', wbFRTOP As Workbook
    Dim shtPPR As Worksheet, shtPID As Worksheet, shtLPIstow As Worksheet, shtMTpick As Worksheet
    Dim shtFRR As Worksheet, shtLPIpick As Worksheet, shtUR As Worksheet ', shtFRTOP As Worksheet
    Dim orow As Integer
    Dim data_path As Variant
    Dim aStrings(1 To 7) As String
    
    orow = 20

    Call InitializeVariables

    aStrings(1) = "data1": aStrings(2) = "data2": aStrings(3) = "data3": _
    aStrings(4) = "data4": aStrings(5) = "data5": aStrings(6) = "data6": _
    aStrings(7) = "data7"


    For Each data_path In aStrings
        'IMPORTANT: Replace the "..." with below by the path to the data files
        folder_path = "C:\Users\geomonto\Desktop\Shift_Comparison_Reporting\" & data_path & "\"
        
        'define all the inputs
        PPR_extension = "PPR.csv"
        'PPR_extension = "processPathReport-KRB1-Day*"
        PPR_input_file = Dir(folder_path & PPR_extension)
        
        PID_extension = "PID.csv"
        'PID_extension = "dat*"
        PID_input_file = Dir(folder_path & PID_extension)
        
        LPIstow_extension = "LPIstow.csv"
        'LPIstow_extension = "processInspector-BIN_USAGE-CONTAINER_TYPE"
        LPIstow_input_file = Dir(folder_path & LPIstow_extension)

        'MTpick_extension = "MTpick.csv"
        'MTpick_input_file = Dir(folder_path & MTpick_extension)

        FRR_extension = "FRR.csv"
        'FRR_extension = "functionRollupReport-KRB1-Transf*"
        FRR_input_file = Dir(folder_path & FRR_extension)

        LPIpick_extension = "LPIpick.csv"
        'LPIpick_extension = "processInspector-CONTAINER_TYPE-GL_CODE"
        LPIpick_input_file = Dir(folder_path & LPIpick_extension)
        
        UR_extension = "UR.csv"
        'UR_extension = "unitsRollup-KRB1-ItemPicked*"
        UR_input_file = Dir(folder_path & UR_extension)
        
        'FRTOP_extension = "FRTOP.csv"
        'FRTOP_input_file = Dir(folder_path & FRTOP_extension)
       
        
        'open workbook
        Set wbPPR = Workbooks.Open(Filename:=folder_path & PPR_input_file)
        Set shtPPR = wbPPR.Sheets(1)
        
        Set wbPID = Workbooks.Open(Filename:=folder_path & PID_input_file)
        Set shtPID = wbPID.Sheets(1)
        
        Set wbLPIstow = Workbooks.Open(Filename:=folder_path & LPIstow_input_file)
        Set shtLPIstow = wbLPIstow.Sheets(1)

        'Set wbMTpick = Workbooks.Open(Filename:=folder_path & MTpick_input_file)
        'Set shtMTpick = wbMTpick.Sheets("MTpick")

        Set wbFRR = Workbooks.Open(Filename:=folder_path & FRR_input_file)
        Set shtFRR = wbFRR.Sheets(1)

        Set wbLPIpick = Workbooks.Open(Filename:=folder_path & LPIpick_input_file)
        Set shtLPIpick = wbLPIpick.Sheets(1)
        
        Set wbUR = Workbooks.Open(Filename:=folder_path & UR_input_file)
        Set shtUR = wbUR.Sheets(1)
        
        'Set wbFRTOP = Workbooks.Open(Filename:=folder_path & FRTOP_input_file)
        'Set shtFRTOP = wbFRTOP.Sheets("FRTOP")
       
        
        
        'Get data
        Call PPR_transfer(shtPPR, shtMain, orow)
        Call PID_transfer(shtPID, shtMain, orow)
        Call LPI_transfer(shtLPIstow, shtMain, 9, orow)
        'Call MTpick_transfer(shtMTpick, shtMain, orow)
        Call LPI_transfer(shtLPIpick, shtMain, 16, orow)
        Call IBCPLH(shtPPR, shtMain, orow)
        Call FRR_transfer(shtFRR, shtMain, 13, orow)
        Call OBCPLH(shtPPR, shtUR, shtMain, 12, orow)
        Call pickRate(shtFRR, shtMain, 10, orow)
        
        
        
        
        'add in 1 row increments
        orow = orow + 1
        
        wbPPR.Close
        wbPID.Close
        wbLPIstow.Close
        'wbMTpick.Close
        wbFRR.Close
        wbLPIpick.Close
        wbUR.Close
        'wbFRTOP.Close
        input_file = Dir
        
    Next data_path
    
Application.ScreenUpdating = True

shtMaint.Activate
End Sub

Function do_something(ByRef sInput As String)

    Debug.Print sInput

End Function

Public Sub PPR_transfer(input_sheet As Worksheet, output_sheet As Worksheet, output_row As Integer)

'Get receive dock values (example J2 = (2,10))
output_sheet.Cells(output_row, 2).value = Round(input_sheet.Cells(2, 10), 1)
'Get stow
output_sheet.Cells(output_row, 4).value = Round(input_sheet.Cells(46, 10), 1)
'Get IB Total
output_sheet.Cells(output_row, 5).value = Round(input_sheet.Cells(54, 10), 1)
'Get Receive Volume
output_sheet.Cells(output_row, 6).value = Round(input_sheet.Cells(54, 8), 1)
'Get inbound UPC
output_sheet.Cells(output_row, 8).value = Round(input_sheet.Cells(54, 8) / input_sheet.Cells(14, 8), 1)
'Get Pick Volume
output_sheet.Cells(output_row, 11).value = Round(input_sheet.Cells(69, 8), 1)
'Get TO Dock
output_sheet.Cells(output_row, 14).value = Round(input_sheet.Cells(71, 10), 1)
'TO total
output_sheet.Cells(output_row, 15).value = Round(input_sheet.Cells(74, 10), 1)

End Sub

Public Sub PID_transfer(input_sheet As Worksheet, output_sheet As Worksheet, output_row As Integer)
'Get LP Receive
output_sheet.Cells(output_row, 3).value = Round(input_sheet.Cells(5, 2), 1)

End Sub

Public Sub LPI_transfer(input_sheet As Worksheet, output_sheet As Worksheet, col As Integer, output_row As Integer)
'Get TOT%
output_sheet.Cells(output_row, col).value = Round(Application.WorksheetFunction.Sum(input_sheet.Range("G:G")) / Application.WorksheetFunction.Sum(input_sheet.Range("H:H")) * 100, 1)

End Sub

Public Sub MTpick_transfer(input_sheet As Worksheet, output_sheet As Worksheet, output_row As Integer)

output_sheet.Cells(output_row, 9).value = Round(input_sheet.Cells(5, 2), 1)

End Sub

Public Sub IBCPLH(input_sheet As Worksheet, output_sheet As Worksheet, output_row As Integer)
'Get IB case per labor hour
output_sheet.Cells(output_row, 7).value = Round(input_sheet.Cells(46, 8) / input_sheet.Cells(180, 9), 1)
End Sub


Public Sub FRR_transfer(input_sheet As Worksheet, output_sheet As Worksheet, col As Integer, output_row As Integer)
'Get Outbound_UPC
output_sheet.Cells(output_row, col).value = Round(Application.SumIfs(input_sheet.Columns(17), input_sheet.Columns(15), "EACH", input_sheet.Columns(16), "Total") / _
Application.SumIfs(input_sheet.Columns(13), input_sheet.Columns(15), "EACH", input_sheet.Columns(16), "Total"), 1)
End Sub

Public Sub OBCPLH(input_sheet1 As Worksheet, input_sheet2 As Worksheet, output_sheet As Worksheet, col As Integer, output_row As Integer)
'Gets pick rate from calucaltion of two numbers from two files
output_sheet.Cells(output_row, col).value = Round(Application.SumIfs(input_sheet2.Columns(9), input_sheet2.Columns(8), "Total", input_sheet2.Columns(7), "Case") / input_sheet1.Cells(181, 9), 1)

End Sub

Public Sub pickRate(input_sheet As Worksheet, output_sheet As Worksheet, col As Integer, output_row As Integer)
'Get pick rate from FRTOP
output_sheet.Cells(output_row, col).value = Round(Application.SumIfs(input_sheet.Columns(17), input_sheet.Columns(16), "Total", input_sheet.Columns(15), "Case") / _
Application.SumIfs(input_sheet.Columns(11), input_sheet.Columns(16), "Total", input_sheet.Columns(15), "Case"), 1)

End Sub




