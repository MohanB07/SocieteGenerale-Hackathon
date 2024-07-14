Option Explicit

Sub GenerateReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim btnGenerate As Shape
    
    ' Set references to worksheets
    Set wsData = ThisWorkbook.Sheets("Data")
    
    ' Check if "Report" sheet exists, create if it doesn't
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Report")
    On Error GoTo 0
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "Report"
    Else
        ' Clear existing report data and controls
        wsReport.Cells.Clear
        DeleteShapes wsReport
    End If
    
    ' Add Generate Report button if not already added
    On Error Resume Next
    Set btnGenerate = wsReport.Shapes("btnGenerate")
    On Error GoTo 0
    If btnGenerate Is Nothing Then
        Set btnGenerate = wsReport.Shapes.AddShape(msoShapeRectangle, 20, 20, 120, 30)
        With btnGenerate
            .TextFrame.Characters.Text = "Generate Report"
            .Name = "btnGenerate"
            .OnAction = "GenerateReport"
            .Fill.ForeColor.RGB = RGB(0, 176, 240) ' Blue fill color
            .TextFrame.Characters.Font.Size = 12
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255) ' White font color
            .TextFrame.HorizontalAlignment = xlHAlignCenter
        End With
    End If
    
    ' Find the last row with data in column A
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Copy the data to the report sheet
    wsData.Range("A1:F" & lastRow).Copy Destination:=wsReport.Range("A1")
    
    ' Generate a chart
    Set chartObj = wsReport.ChartObjects.Add(160, 60, 600, 300)
    With chartObj.Chart
        .SetSourceData Source:=wsReport.Range("A1:F" & lastRow)
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Financial Data Overview"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Value"
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = False ' Hide gridlines
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
    End With
    
    ' Add comparison options buttons
    AddComparisonButtons wsReport
End Sub

Sub AddComparisonButtons(wsReport As Worksheet)
    Dim btnCompareDay As Shape
    Dim btnCompareWeek As Shape
    Dim btnCompareMonth As Shape
    Dim btnCustomRange As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    
    ' Button dimensions and positions
    leftPos = 160
    topPos = 400
    btnWidth = 120
    btnHeight = 30
    
    ' Create Compare Day button
    Set btnCompareDay = wsReport.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, btnWidth, btnHeight)
    With btnCompareDay
        .TextFrame.Characters.Text = "Compare Day"
        .Name = "btnCompareDay"
        .OnAction = "CompareDay"
        .Fill.ForeColor.RGB = RGB(255, 192, 0) ' Yellow fill color
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0) ' Black font color
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    ' Create Compare Week button
    leftPos = leftPos + btnWidth + 20
    Set btnCompareWeek = wsReport.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, btnWidth, btnHeight)
    With btnCompareWeek
        .TextFrame.Characters.Text = "Compare Week"
        .Name = "btnCompareWeek"
        .OnAction = "CompareWeek"
        .Fill.ForeColor.RGB = RGB(0, 176, 240) ' Blue fill color
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255) ' White font color
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    ' Create Compare Month button
    leftPos = leftPos + btnWidth + 20
    Set btnCompareMonth = wsReport.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, btnWidth, btnHeight)
    With btnCompareMonth
        .TextFrame.Characters.Text = "Compare Month"
        .Name = "btnCompareMonth"
        .OnAction = "CompareMonth"
        .Fill.ForeColor.RGB = RGB(146, 208, 80) ' Green fill color
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0) ' Black font color
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
    ' Create Custom Range button
    leftPos = leftPos + btnWidth + 20
    Set btnCustomRange = wsReport.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, btnWidth, btnHeight)
    With btnCustomRange
        .TextFrame.Characters.Text = "Custom Range"
        .Name = "btnCustomRange"
        .OnAction = "CustomRange"
        .Fill.ForeColor.RGB = RGB(255, 0, 0) ' Red fill color
        .TextFrame.Characters.Font.Size = 12
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255) ' White font color
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
End Sub

Sub DeleteShapes(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
End Sub

Sub CompareDay()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' Find the last row with data in column A
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Get data range for the current and previous day
    If lastRow >= 2 Then
        Set dataRange = wsData.Range("A" & lastRow - 1 & ":F" & lastRow)
    Else
        MsgBox "Not enough data for comparison."
        Exit Sub
    End If
    
    ' Clear previous comparison data
    wsReport.Cells.Clear
    
    ' Copy the data to the report sheet
    dataRange.Copy Destination:=wsReport.Range("A1")
    
    ' Generate a chart
    Set chartObj = wsReport.ChartObjects.Add(160, 60, 600, 300)
    With chartObj.Chart
        .SetSourceData Source:=wsReport.Range("A1:F2")
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Comparison: Current vs Previous Day"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Value"
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = False ' Hide gridlines
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
    End With
End Sub

Sub CompareWeek()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' Find the last row with data in column A
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Get data range for the current and previous week
    If lastRow >= 7 Then
        Set dataRange = wsData.Range("A" & lastRow - 6 & ":F" & lastRow)
    Else
        MsgBox "Not enough data for comparison."
        Exit Sub
    End If
    
    ' Clear previous comparison data
    wsReport.Cells.Clear
    
    ' Copy the data to the report sheet
    dataRange.Copy Destination:=wsReport.Range("A1")
    
    ' Generate a chart
    Set chartObj = wsReport.ChartObjects.Add(160, 60, 600, 300)
    With chartObj.Chart
        .SetSourceData Source:=wsReport.Range("A1:F7")
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Comparison: Current vs Previous Week"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Value"
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = False ' Hide gridlines
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
    End With
End Sub

Sub CompareMonth()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' Find the last row with data in column A
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    ' Get data range for the current and previous month
    If lastRow >= 30 Then
        Set dataRange = wsData.Range("A" & lastRow - 29 & ":F" & lastRow)
    Else
        MsgBox "Not enough data for comparison."
        Exit Sub
    End If
    
    ' Clear previous comparison data
    wsReport.Cells.Clear
    
    ' Copy the data to the report sheet
    dataRange.Copy Destination:=wsReport.Range("A1")
    
    ' Generate a chart
    Set chartObj = wsReport.ChartObjects.Add(160, 60, 600, 300)
    With chartObj.Chart
        .SetSourceData Source:=wsReport.Range("A1:F30")
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Comparison: Current vs Previous Month"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Value"
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = False ' Hide gridlines
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
    End With
End Sub

Sub CustomRange()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim chartObj As ChartObject
    Dim startDate As Date
    Dim endDate As Date
    Dim dataRange As Range
    Dim filteredRange As Range
    Dim copyRange As Range
    Dim destRange As Range
    Dim rowCount As Long
    Dim dateFormat As String
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' Set the date format expected from the input box
    dateFormat = "yyyy-mm-dd"
    
    ' Prompt user for custom range
    On Error Resume Next
    startDate = DateValue(InputBox("Enter start date (format: YYYY-MM-DD)", "Start Date"))
    endDate = DateValue(InputBox("Enter end date (format: YYYY-MM-DD)", "End Date"))
    On Error GoTo 0
    
    If startDate = 0 Or endDate = 0 Then
        MsgBox "Invalid date format. Please enter dates in format: YYYY-MM-DD"
        Exit Sub
    End If
    
    ' Find data within custom range
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    Set dataRange = wsData.Range("A1:F" & lastRow)
    
    ' Filter data for the custom range
    dataRange.AutoFilter Field:=1, Criteria1:=">=" & Format(startDate, dateFormat), Operator:=xlAnd, Criteria2:="<=" & Format(endDate, dateFormat)
    
    ' Check if any data meets the criteria
    On Error Resume Next
    Set filteredRange = dataRange.Offset(1, 0).Resize(dataRange.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If filteredRange Is Nothing Then
        MsgBox "No data found for the selected date range."
        wsData.AutoFilterMode = False ' Remove the filter
        Exit Sub
    End If
    
    ' Copy the filtered data to the report sheet
    rowCount = filteredRange.Rows.Count
    Set copyRange = filteredRange.Resize(, 6) ' Ensure to copy all columns A to F
    Set destRange = wsReport.Range("A1").Resize(rowCount, 6)
    
    copyRange.Copy Destination:=destRange
    
    ' Remove the filter
    wsData.AutoFilterMode = False
    
    ' Generate a chart
    Set chartObj = wsReport.ChartObjects.Add(160, 60, 600, 300)
    With chartObj.Chart
        .SetSourceData Source:=wsReport.Range("A1:F" & rowCount)
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Custom Range Comparison"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Date"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Value"
        .Axes(xlCategory).TickLabelPosition = xlLow
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = False ' Hide gridlines
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
    End With
End Sub

