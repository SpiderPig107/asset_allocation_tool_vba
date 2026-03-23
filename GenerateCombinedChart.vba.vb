Attribute VB_Name = "GenerateCombinedChart"
Sub GenerateCombinedStockCharts()

Dim ws_control_panel As Worksheet
Dim stock1 As String
Dim stock2 As String
Dim date_row_length_stock1 As Long
Dim date_row_length_stock2 As Long
Dim stock1_chart As ChartObject
Dim stock2_chart As ChartObject

    ' 1. Identify the first company from Control Panel
    Set ws_control_panel = ThisWorkbook.Sheets("Control_Panel")
    stock1 = ws_control_panel.range("B2").Value
    stock2 = ws_control_panel.range("B3").Value
    
    ' 2. Create a sheet to display the charts
    
    ' 2.1. If there are old charts, this deletes that so the sheet is new
    On Error Resume Next                    ' If it encounters error, just continue on
    Application.DisplayAlerts = False   ' Disables any alert notification (e.g. Msgbox)
    Sheets("Combined_Stock_Report").Delete
    Application.DisplayAlerts = True    ' Re-enable alert notification so Excel will notify you if you want to delete a sheet
    On Error GoTo 0                             ' This tells Excel to go back to original state where if there is an error, it will stop
    
    ' 2.2. Create a new sheet called "Combined_Stock_Report"
    Set ws_combined_stock_report = ThisWorkbook.Sheets.Add(After:=Worksheets("Control_Panel"))
    ws_combined_stock_report.Name = "Combined_Stock_Report"
    
    ' 3. Measure the row length of dates from the individual stock sheet
    date_row_length_stock1 = Worksheets(stock1).Cells(Rows.Count, 1).End(xlUp).Row
    date_row_length_stock2 = Worksheets(stock2).Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Rows.Count is VBA-way of saying for "the very last row possible". This is like "Ctrl + Down". 1 represents the first column so Column A.
    ' End(xlUp) It is the VBA equivalent of pressing Ctrl + Up Arrow on your keyboard.
    ' Without .Row at the end, VBA wouldn't give you a number; it would give you the entire cell.
    
    ' 4. Copy the Dates and Closing Price of both stocks to "Combined_Stock_Report" sheet
    With ws_combined_stock_report.range("A1")
        .Value = "Date"
        .Font.Bold = True
    End With
    
    With ws_combined_stock_report.range("B1")
        .Value = stock1 & " (Closing Price)"
        .Font.Bold = True
    End With
    
    Worksheets(stock1).range("A2:A" & date_row_length_stock1).Copy ws_combined_stock_report.range("A2") ' Copy and paste "Date" column
    Worksheets(stock1).range("E2:E" & date_row_length_stock1).Copy ws_combined_stock_report.range("B2") ' Copy and paste "Close" column

    With ws_combined_stock_report.range("E1")
        .Value = "Date"
        .Font.Bold = True
    End With
    
    With ws_combined_stock_report.range("F1")
        .Value = stock2 & " (Closing Price)"
        .Font.Bold = True
    End With
    
    Worksheets(stock2).range("A2:A" & date_row_length_stock2).Copy ws_combined_stock_report.range("E2") ' Copy and paste "Date" column
    Worksheets(stock2).range("E2:E" & date_row_length_stock2).Copy ws_combined_stock_report.range("F2") ' Copy and paste "Close" column
    
    ' 4. Create an empty chart
    Set combined_chart = Worksheets("Combined_Stock_Report").ChartObjects.Add( _
        Top:=Worksheets("Combined_Stock_Report").Cells(1, 9).Top, _
        Left:=Worksheets("Combined_Stock_Report").Cells(1, 9).Left, _
        Width:=600, Height:=500)
    
    ' " _" is to tell excel that the code is too long to fit in a single line, the following line is a continuation of it.
    ' Cells(1, 15) is a location on a grid, but .Top translates that location into a measurement of distance.
    ' you are asking VBA to do a conversion:
    ' The Input: A grid coordinate (Row 1, Column 15).
    ' The Output: A physical distance from the top of the screen (e.g., 0 or 15.75).
    ' Top refers to the Vertical distance (think of it like y-axis)
    ' Left refers to the Horizontal distance (think of it like x-axis)
    
    ' 5. Link the data range to the empty chart
    With combined_chart.Chart
        .SetSourceData Source:=Worksheets("Combined_Stock_Report").range("A1:A" & date_row_length_stock1 & ",B1:B" & date_row_length_stock1 & ",F1:F" & date_row_length_stock1)
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Price Trend for " & stock1 & " and " & stock2
        .Axes(xlValue).HasMajorGridlines = False    ' Remove gridlines
    End With
    
    Call SavingsPerformanceChart.SavingsChart_Both
    
End Sub

