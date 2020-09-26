Attribute VB_Name = "website_opener"
Sub openWebsites()
  OpenInFireFoxNewTab "https://fclm-portal.amazon.com/reports/processPathRollup?reportFormat=HTML&warehouseId=KRB1&spanType=Day&startDateDay=2020%2F09%2F13&maxIntradayDays=1&startHourIntraday=0&startMinuteIntraday=0&endHourIntraday=0&endMinuteIntraday=0&_adjustPlanHours=on&_hideEmptyLineItems=on&employmentType=AllEmployees"
  WasteTime (0.5)
  OpenInFireFoxNewTab "https://monitorportal.amazon.com/igraph?SchemaName1=Service&DataSet1=Prod&Marketplace1=KRB1&HostGroup1=ALL&Host1=ALL&ServiceName1=AFTInboundDirectorService&MethodName1=PerformanceHealthHandler&Client1=ALL&MetricClass1=PID&Instance1=PID-1&Metric1=Encounter.FinalState.RECEIVED&Period1=OneHour&Stat1=sum&Label1=Encounter.FinalState.RECEIVED&SchemaName2=Service&Metric2=Encounter.FinalState.CANNOT_CHECK_IN&Label2=Encounter.FinalState.CANNOT_CHECK_IN&SchemaName3=Service&Metric3=Encounter.FinalState.CANNOT_RECEIVE&Label3=Encounter.FinalState.CANNOT_RECEIVE&HeightInPixels=250&WidthInPixels=600&GraphTitle=KRB1%20PID-1&DecoratePoints=true&StartTime1=2020-09-13T14%3A00%3A00Z&EndTime1=2020-09-14T01%3A00%3A00Z&FunctionExpression1=SUM%28M1%2CM2%2CM3%29&FunctionLabel1=AVG%20%28avg%3A%20%7Bavg%7D%29&FunctionYAxisPreference1=left&FunctionColor1=default"
  WasteTime (0.5)
  OpenInFireFoxNewTab "https://fclm-portal.amazon.com/ppa/inspect/process?primaryAttribute=BIN_USAGE&secondaryAttribute=CONTAINER_TYPE&nodeType=FC&warehouseId=KRB1&processId=100004&spanType=Day&startDateDay=2020%2F09%2F09&startDateWeek=2020%2F08%2F14&startDateMonth=2020%2F08%2F01&maxIntradayDays=1&startDateIntraday=2020%2F08%2F14&startHourIntraday=0&startMinuteIntraday=0&endDateIntraday=2020%2F08%2F14&endHourIntraday=0&endMinuteIntraday=0"
  WasteTime (0.5)
  OpenInFireFoxNewTab "https://fclm-portal.amazon.com/reports/functionRollup?reportFormat=HTML&warehouseId=KRB1&processId=1003065&spanType=Day&startDateDay=2020%2F09%2F13&maxIntradayDays=1&startHourIntraday=0&startMinuteIntraday=0&endHourIntraday=0&endMinuteIntraday=0"
  WasteTime (0.5)
  OpenInFireFoxNewTab "https://fclm-portal.amazon.com/ppa/inspect/process?primaryAttribute=CONTAINER_TYPE&secondaryAttribute=GL_CODE&nodeType=FC&warehouseId=KRB1&processId=100115&spanType=Day&startDateDay=2020%2F09%2F09&startDateWeek=2020%2F08%2F14&startDateMonth=2020%2F08%2F01&maxIntradayDays=1&startDateIntraday=2020%2F08%2F14&startHourIntraday=0&startMinuteIntraday=0&endDateIntraday=2020%2F08%2F14&endHourIntraday=0&endMinuteIntraday=0"
  WasteTime (0.5)
  OpenInFireFoxNewTab "https://fclm-portal.amazon.com/reports/unitsRollup?reportFormat=HTML&warehouseId=KRB1&jobAction=ItemPicked&startDate=2020%2F09%2F13&startHour=7&startMinute=0&endDate=2020%2F09%2F13&endHour=18&endMinute=0"
End Sub

Sub OpenInFireFoxNewTab(url As String)
  Dim pathFireFox As String
  pathFireFox = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
  If Dir(pathFireFox) = "" Then pathFireFox = "C:\Program Files\Mozilla Firefox\firefox.exe"
  If Dir(pathFireFox) = "" Then
    MsgBox "FireFox Path Not Found", vbCritical, "Macro Ending"
    Exit Sub
  End If
  Shell """" & pathFireFox & """" & " -new-tab " & url, vbHide
   
End Sub

Sub WasteTime(Finish As Long)
 
    Dim NowTick As Long
    Dim EndTick As Long
 
    EndTick = GetTickCount + (Finish * 1000)
     
    Do
 
        NowTick = GetTickCount
        DoEvents
 
    Loop Until NowTick >= EndTick
 
End Sub
