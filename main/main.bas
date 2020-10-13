Attribute VB_Name = "main"
Option Explicit

Sub main()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False

    Dim dtDate As Date, dtStartDate As Date, dtEndDate As Date
    Dim iter As Integer
    Dim stIter As String
    Dim building As String


    dtStartDate = Range("B2").value '#1/1/2018#
    dtEndDate = dtStartDate + 3
    

    
    building = Range("B3").value
    
    iter = 0

    For dtDate = dtStartDate To dtEndDate
        iter = iter + 1
        stIter = CStr(iter)
        
        '!!!!!!!!!!!!!!! ToDo: array and loop for database names !!!!!!!!!
        Call import("ppr", stIter, dtDate, building)
        Call import("pid", stIter, dtDate, building)
        Call import("frr", stIter, dtDate, building)
        Call import("ur", stIter, dtDate, building)
        
        
        'Excel is horrible so feed it slow
        Application.Wait (Now + TimeValue("0:00:01"))
    'Debug.Print (dtDate)
        
    Next dtDate
    
    'Call delayedSort
    
    Application.ScreenUpdating = True
    Sheets("Report Generator").Range("D2").Select

End Sub


Sub import(dataBase As String, refIter As String, dtDate, building)
'''' THIS SUB MAKES SURE THE RIGHT WORKSHEETS ARE PRESENT OR CREATES THEM'''
    Dim Flag
    Dim Count
    Dim i
    Dim wsName

    'name of worksheet iteration"
    'refIter = dataBase + refIter
    'Debug.Print dataBase
    
    Flag = 0
    Count = ActiveWorkbook.Worksheets.Count
    
        For i = 1 To Count
        
            wsName = ActiveWorkbook.Worksheets(i).Name
            If wsName = dataBase + refIter Then Flag = 1
            'If wsName = refIter Then Flag = 1
            
        Next i
        
            If Flag = 1 Then
                Debug.Print dataBase & refIter & " worksheet exist."
            Else
                Debug.Print dataBase & refIter & " worksheet was created"
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = dataBase + refIter
            End If
            
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Call websiteDictionary(dataBase, refIter, dtDate, building)
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    dataBase = vbNullString
    
    Sheets("Report Generator").Select

    Debug.Print "Connecting to import data for " & dtDate & " ..."
    
End Sub


Sub sort()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    
''' THIS SUB TRANSFORMS ARRAYS TO COLUMN/CELL FORMAT AND MAPS DATA ONTO REPORT'''
'!!!!!!!!!!!!!! FIX: horrible way of doing this (e.g. loop through the rows)
    Dim itemm As Worksheet
    Dim arrWs
    
    'Set arrWs = Sheets(Array("ppr1", "ppr2", "ppr3", "ppr4", "ppr5", "ppr6", "ppr7", _
    '                         "pid1", "pid2", "pid3", "pid4", "pid5", "pid6", "pid7" _
    '                        ))
                            
    Set arrWs = Sheets(Array("ppr1", "ppr2", "ppr3", "ppr4", _
                             "pid1", "pid2", "pid3", "pid4", _
                             "frr1", "frr2", "frr3", "frr4", _
                             "ur1", "ur2", "ur3", "ur4" _
                            ))

    For Each itemm In arrWs
        itemm.Select
        Columns("A:A").Select
        Call transform
        itemm.Columns.AutoFit
        If itemm.Name = "ppr1" Then
            Call mapPPR(itemm, 14)
        ElseIf itemm.Name = "ppr2" Then
            Call mapPPR(itemm, 15)
        ElseIf itemm.Name = "ppr3" Then
            Call mapPPR(itemm, 16)
        ElseIf itemm.Name = "ppr4" Then
            Call mapPPR(itemm, 17)
        ElseIf itemm.Name = "ppr5" Then
            Call mapPPR(itemm, 18)
        ElseIf itemm.Name = "ppr6" Then
            Call mapPPR(itemm, 19)
        ElseIf itemm.Name = "ppr7" Then
            Call mapPPR(itemm, 20)
            
        ElseIf itemm.Name = "pid1" Then
            Call mapPID(itemm, 14)
        ElseIf itemm.Name = "pid2" Then
            Call mapPID(itemm, 15)
        ElseIf itemm.Name = "pid3" Then
            Call mapPID(itemm, 16)
        ElseIf itemm.Name = "pid4" Then
            Call mapPID(itemm, 17)
        ElseIf itemm.Name = "pid5" Then
            Call mapPID(itemm, 18)
        ElseIf itemm.Name = "pid6" Then
            Call mapPID(itemm, 19)
        ElseIf itemm.Name = "pid7" Then
            Call mapPID(itemm, 20)
            
        ElseIf itemm.Name = "frr1" Then
            Call mapFRR(itemm, 14)
        ElseIf itemm.Name = "frr2" Then
            Call mapFRR(itemm, 15)
        ElseIf itemm.Name = "frr3" Then
            Call mapFRR(itemm, 16)
        ElseIf itemm.Name = "frr4" Then
            Call mapFRR(itemm, 17)
        ElseIf itemm.Name = "frr5" Then
            Call mapFRR(itemm, 18)
        ElseIf itemm.Name = "frr6" Then
            Call mapFRR(itemm, 19)
        ElseIf itemm.Name = "frr7" Then
            Call mapFRR(itemm, 20)
            
        ElseIf itemm.Name = "ur1" Then
            Call mapUR(itemm, Sheets("ppr1"), 14)
        ElseIf itemm.Name = "ur2" Then
            Call mapUR(itemm, Sheets("ppr2"), 15)
        ElseIf itemm.Name = "ur3" Then
            Call mapUR(itemm, Sheets("ppr3"), 16)
        ElseIf itemm.Name = "ur4" Then
            Call mapUR(itemm, Sheets("ppr4"), 17)
        ElseIf itemm.Name = "ur5" Then
            Call mapUR(itemm, Sheets("ppr5"), 18)
        ElseIf itemm.Name = "ur6" Then
            Call mapUR(itemm, Sheets("ppr6"), 19)
        ElseIf itemm.Name = "ur7" Then
            Call mapUR(itemm, Sheets("ppr7"), 20)
        Else
            Debug.Print itemm, "woksheet does not exist"
        End If

    Next itemm
    
    Application.ScreenUpdating = True
    
  
Sheets("Report Generator").Select

End Sub


Sub mapPPR(ws As Worksheet, j As Integer)
    
    '''''map data onto report
        'Get reveive dock values
        Sheets("Report Generator").Cells(j, 2).value = Round(ws.Cells(2, 10), 1)
        'Get stow
        Sheets("Report Generator").Cells(j, 4).value = Round(ws.Cells(46, 10), 1)
        'Get IB Total
        Sheets("Report Generator").Cells(j, 5).value = Round(ws.Cells(54, 10), 1)
        'Get Receive Volume
        Sheets("Report Generator").Cells(j, 6).value = Round(ws.Cells(54, 8), 1)
        'Get inbound UPC
        Sheets("Report Generator").Cells(j, 8).value = Round(ws.Cells(54, 8) / ws.Cells(14, 8), 1)
        'Get Pick Volume
        Sheets("Report Generator").Cells(j, 11).value = Round(ws.Cells(69, 8), 1)
        'Get TO Dock
        Sheets("Report Generator").Cells(j, 14).value = Round(ws.Cells(71, 10), 1)
        'TO total
        Sheets("Report Generator").Cells(j, 15).value = Round(ws.Cells(74, 10), 1)
        'Get IB case per labor hour
        Sheets("Report Generator").Cells(j, 7).value = Round(ws.Cells(46, 8) / ws.Cells(180, 9), 1)


End Sub

Sub mapPID(ws As Worksheet, j As Integer)
    '''' map PID data report '''
    'LP receive
    Sheets("Report Generator").Cells(j, 3).value = Round(ws.Cells(5, 2), 1)

End Sub


Sub mapFRR(ws As Worksheet, j As Integer)
    
    'gets pick rate
    Sheets("Report Generator").Cells(j, 10).value = Round(Application.SumIfs(ws.Columns(17), ws.Columns(16), "Total", ws.Columns(15), "Case") / _
    Application.SumIfs(ws.Columns(11), ws.Columns(16), "Total", ws.Columns(15), "Case"), 1)
    
    'Outbound UPC
    Sheets("Report Generator").Cells(j, 13).value = Round(Application.SumIfs(ws.Columns(17), ws.Columns(15), "EACH", ws.Columns(16), "Total") / _
    Application.SumIfs(ws.Columns(13), ws.Columns(15), "EACH", ws.Columns(16), "Total"), 1)

End Sub


Sub mapUR(ws As Worksheet, ws2 As Worksheet, j As Integer)
    
    'gets OB CLPH from PPR and UR calculation
    Sheets("Report Generator").Cells(j, 12).value = Round(Application.SumIfs(ws.Columns(9), ws.Columns(8), "Total", ws.Columns(7), "Case") / ws2.Range("I181"), 1)

End Sub



Sub transform()

    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1)), TrailingMinusNumbers:=True

End Sub


Sub delayedSort()
'''THIS SUB HELPS DELAY SUB '''
  Application.OnTime Now() + TimeValue("0:00:30"), "sortPPR"
  'Application.Wait (Now + TimeValue("0:00:30")), "sortPPR"
  sortPPR
  Debug.Print "sorting..."

End Sub
