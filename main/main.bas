Attribute VB_Name = "mainProcedure"
Option Explicit

Sub mainProcedure()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False

    Dim dtDate As Date, dtStartDate As Date, dtEndDate As Date
    Dim iter As Integer
    Dim stIter As String
    Dim building As String

    dtStartDate = Range("B2").value '#1/1/2018#
    dtEndDate = dtStartDate + 6
    
    building = Range("B3").value
    
    iter = 0

    For dtDate = dtStartDate To dtEndDate
        iter = iter + 1
        stIter = CStr(iter)
        Call import("ppr", stIter, dtDate, building)

        Application.Wait (Now + TimeValue("0:00:01"))
    'Debug.Print (dtDate)
        
    Next dtDate
    
    'Call delayedSort
    Application.ScreenUpdating = True

End Sub


Sub import(dtBase As String, refIter As String, dtDate, building)
'''' THIS SUB MAKES SURE THE RIGHT WORKSHEETS ARE PRESENT OR CREATES THEM'''
    Dim Flag
    Dim Count
    Dim i
    Dim wsName

    'name of worksheet iteration"
    refIter = dtBase + refIter
    
    Flag = 0
    Count = ActiveWorkbook.Worksheets.Count
    
        For i = 1 To Count
        
            wsName = ActiveWorkbook.Worksheets(i).Name
            If wsName = refIter Then Flag = 1
            
        Next i
        
            If Flag = 1 Then
                Debug.Print refIter & " worksheet exist."
            Else
                Debug.Print refIter & " worksheet was created"
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = refIter
            End If
            
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim startYear As String, startMonth As String, startDay As String
    
    'decompose date for URL input
    startYear = Year(dtDate)
    startMonth = Month(dtDate)
    startDay = Day(dtDate)
    
    Sheets(refIter).Select
    Cells.Select
    Selection.ClearContents
    
    With ActiveSheet.QueryTables.Add(Connection:="URL;https://fclm-portal.amazon.com/reports/processPathRollup?reportFormat=CSV&warehouseId=" & building & "&spanType=Day&startDateDay=" & startYear & "%2F" & startMonth & "%2F" & startDay & "&maxIntradayDays=1&startHourIntraday=0&startMinuteIntraday=0&endHourIntraday=0&endMinuteIntraday=0&_adjustPlanHours=on&_hideEmptyLineItems=on&employmentType=AllEmployees", Destination:=Sheets(refIter).Range("A1"))
    
        .Name = "website" & startDay 'makes sure that it connects to different websites
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebFormatting = xlWebFormattingNone
        .WebTables = "2"
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=True
    End With
    
    Sheets("Report Generator").Select
    'Sheets(refIter).Select

    'Debug.Print startDay
    Debug.Print "Connecting to import data for " & dtDate & " ..."
End Sub

Sub delayedSort()
'''THIS SUB HELPS DELAY SUB '''
  Application.OnTime Now() + TimeValue("0:00:30"), "sortPPR"
  'Application.Wait (Now + TimeValue("0:00:30")), "sortPPR"
  sortPPR
  Debug.Print "sorting..."

End Sub


Sub sortPPR()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    
''' THIS SUB TRANSFORMS ARRAYS TO COLUMN/CELL FORMAT AND MAPS DATA ONTO REPORT'''
    'Dim ppr1, ppr2, ppr3, ppr4, ppr5, ppr6, ppr7 As String
    Dim itemm As Worksheet
    Dim arrWs
    
    Set arrWs = Sheets(Array("ppr1", "ppr2", "ppr3", "ppr4", "ppr5", "ppr6", "ppr7"))

    For Each itemm In arrWs
        itemm.Select
        Columns("A:A").Select
        Call sort
        itemm.Columns.AutoFit
        'MsgBox itemm.Range("a1")
        'Debug.Print itemm.Name
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
        End If

    Next itemm
    
    Application.ScreenUpdating = True
    
  
Sheets(1).Select

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


End Sub


Sub sort()

    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1)), TrailingMinusNumbers:=True

End Sub
