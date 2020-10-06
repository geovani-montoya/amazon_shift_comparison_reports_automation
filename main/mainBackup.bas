Attribute VB_Name = "mainBackup"
Option Explicit

Sub mainProcedure()

    Dim dtDate As Date, dtStartDate As Date, dtEndDate As Date
    Dim iter As Integer
    Dim stIter As String
    Dim building As String

    dtStartDate = Range("A14").value '#1/1/2018#
    dtEndDate = dtStartDate + 2
    
    building = Range("A6").value
    
    iter = 0

    For dtDate = dtStartDate To dtEndDate
        iter = iter + 1
        stIter = CStr(iter)
        Call checkWorksheet(stIter, dtDate, building)
        'Call getPPR_main(dDate, building)
        'Call delayedSort
        Application.Wait (Now + TimeValue("0:00:01"))
    'Debug.Print (dtDate)
        
    Next dtDate

End Sub


Public Sub checkWorksheet(refIter As String, dtDate, building)
'''' THIS SUB MAKES SURE THE RIGHT WORKSHEETS ARE PRESENT OR CREATES THEM'''
    Dim Flag
    Dim Count
    Dim i
    Dim wsName

    'name of worksheet iteration"
    refIter = "ppr" + refIter
    
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
    
    
    Sheets(refIter).Select

    Debug.Print startDay
    Debug.Print "Connecting to import data for " & dtDate
End Sub

