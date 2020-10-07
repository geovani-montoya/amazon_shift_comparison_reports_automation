Attribute VB_Name = "websitesDictionary"
Option Explicit

Public Sub websiteDictionary(dataBase, refIter, dtDate, building)
''' THIS SUB FINDS LETS MAIN USE THE CORRECT LINK TO THE WEBSITE '''

    Dim startYear As String, startMonth As String, startDay As String
    
    'decompose date for URL input
    startYear = Year(dtDate)
    startMonth = Month(dtDate)
    startDay = Day(dtDate)
    
    Sheets(refIter).Select
    Cells.Select
    Selection.ClearContents
    
    If dataBase = "ppr1" Then
    
        With ActiveSheet.QueryTables.Add(Connection:="URL;https://fclm-portal.amazon.com/reports/" _
            & "processPathRollup?reportFormat=CSV&warehouseId=" & building & "&spanType=Day&startDateDay=" _
            & startYear & "%2F" & startMonth & "%2F" & startDay & "&maxIntradayDays=1&startHourIntraday=0" _
            & "&startMinuteIntraday=0&endHourIntraday=0&endMinuteIntraday=0&_adjustPlanHours=on&_hideEmptyLineItems=on" _
            & "&employmentType=AllEmployees", Destination:=Sheets(refIter).Range("A1"))
    
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
        
        
    ElseIf dataBase = "ppr" Then
        
        With ActiveSheet.QueryTables.Add(Connection:="URL;https://monitorportal.amazon.com/mws?Action=" _
        & "GetGraph&Version=2007-07-07&SchemaName1=Service&DataSet1=Prod&Marketplace1=KRB1&HostGroup1=" _
        & "ALL&Host1=ALL&ServiceName1=AFTInboundDirectorService&MethodName1=PerformanceHealthHandler" _
        & "&Client1=ALL&MetricClass1=PID&Instance1=PID-1&Metric1=Encounter.FinalState.RECEIVED&Period1=" _
        & "OneHour&Stat1=sum&Label1=Encounter.FinalState.RECEIVED&SchemaName2=Service&Metric2=" _
        & "Encounter.FinalState.CANNOT_CHECK_IN&Label2=Encounter.FinalState.CANNOT_CHECK_IN&SchemaName3=" _
        & "Service&Metric3=Encounter.FinalState.CANNOT_RECEIVE&Label3=Encounter.FinalState." _
        & "CANNOT_RECEIVE&HeightInPixels=250&WidthInPixels=600&GraphTitle=KRB1%20PID-1&" _
        & "DecoratePoints=true&StartTime1=2020-09-30T14%3A00%3A00Z&EndTime1=2020-10-01T01%3A00%3A00Z&" _
        & "FunctionExpression1=SUM%28M1%2CM2%2CM3%29&FunctionLabel1=AVG%20%28avg%3A%20%7Bavg%7D%29&" _
        & "FunctionYAxisPreference1=left&FunctionColor1=default&OutputFormat=CSV_TRANSPOSE" _
        , Destination:=Sheets(refIter).Range("A1"))
        
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

        
    End If
    




End Sub
