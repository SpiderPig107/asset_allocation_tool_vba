Attribute VB_Name = "GenerateIndividualStockCharts"
Sub GenerateStockCharts()

Dim ws_control_panel As Worksheet
Dim stock1 As String
Dim stock2 As String
Dim date_row_length_stock1 As Long
Dim date_row_length_stock2 As Long
Dim stock1_chart As ChartObject
Dim stock2_chart As ChartObject

    ' 1. Identify the two company from Control Panel
    Set ws_control_panel = ThisWorkbook.Sheets("Control_Panel")
    stock1 = ws_control_panel.range("B2").Value
    stock2 = ws_control_panel.range("B3").Value
    
    ' 2. Create a sheet to display the charts
    
    ' 2.1. If there are old charts, this deletes that so the sheet is new
    On Error Resume Next                    ' If it encounters error, just continue on
    Application.DisplayAlerts = False   ' Disables any alert notification (e.g. Msgbox)
    Sheets("Combined_Stock_Report").Delete
    Sheets("Stock_Report").Delete
    Application.DisplayAlerts = True    ' Re-enable alert notification so Excel will notify you if you want to delete a sheet
    On Error GoTo 0                             ' This tells Excel to go back to original state where if there is an error, it will stop
    
    ' 2.2. Create a new sheet called "Stock Report"
    Set ws_stock_report = ThisWorkbook.Sheets.Add(After:=Worksheets("Control_Panel"))
    ws_stock_report.Name = "Stock_Report"
    
    ' 3. Measure the row length of dates from the individual stock sheet
    date_row_length_stock1 = Worksheets(stock1).Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Rows.Count is VBA-way of saying for "the very last row possible". This is like "Ctrl + Down". 1 represents the first column so Column A.
    ' End(xlUp) It is the VBA equivalent of pressing Ctrl + Up Arrow on your keyboard.
    ' Without .Row at the end, VBA wouldn't give you a number; it would give you the entire cell.
    
    ' 4. Create an empty chart (stock1)
    Set stock1_chart = Worksheets("Stock_Report").ChartObjects.Add( _
        Top:=Worksheets("Stock_Report").Cells(1, 1).Top, _
        Left:=Worksheets("Stock_Report").Cells(1, 1).Left, _
        Width:=600, Height:=500)
        
    ' " _" is to tell excel that the code is too long to fit in a single line, the following line is a continuation of it.
    ' Cells(1, 15) is a location on a grid, but .Top translates that location into a measurement of distance.
    ' you are asking VBA to do a conversion:
    ' The Input: A grid coordinate (Row 1, Column 15).
    ' The Output: A physical distance from the top of the screen (e.g., 0 or 15.75).
    ' Top refers to the Vertical distance (think of it like y-axis)
    ' Left refers to the Horizontal distance (think of it like x-axis)
    
    ' 5.1. Link the data range to the empty chart (stock1)
    With stock1_chart.Chart
        .SetSourceData Source:=Worksheets(stock1).range("A1:A" & date_row_length_stock1 & " ,E1:E" & date_row_length_stock1)
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Price Trend for " & stock1
        .Axes(xlValue).HasMajorGridlines = False    ' Remove gridlines
    End With
    
    ' 5.2. Add the data for stock 1 below the chart for stock 1
    ' ----Date Column: Stock 1----
    With Worksheets("Stock_Report").range("A35")
        .Value = "Date"
        .Font.Bold = True
    End With
    
    Worksheets(stock1).range("A2:A" & date_row_length_stock1).Copy Worksheets("Stock_Report").range("A36") ' Copy and paste "Date" column
    
    ' ----Closing Price Column: Stock 1----
    With Worksheets("Stock_Report").range("B35")
        .Value = stock1 & " (Closing Price)"
        .Font.Bold = True
    End With
    
    Worksheets(stock1).range("E2:E" & date_row_length_stock1).Copy Worksheets("Stock_Report").range("B36") ' Copy and paste "Close" column
    
    ' ----Investment Value Column: Stock 1----
    Dim initial_value_stock1 As Double
    Dim yesterday_value_stock1 As Double
    Dim current_price_stock1 As Double
    Dim yesterday_price_stock1 As Double
    
    initial_value_stock1 = Worksheets("Control_Panel").range("D2").Value
    
    With Worksheets("Stock_Report").range("C35")
        .Value = "Investment Value "
        .Font.Bold = True
    End With
    
    Worksheets("Stock_Report").range("C36").Value = initial_value_stock1
    
    For i = 2 To date_row_length_stock1
    
        yesterday_value_stock1 = Worksheets("Stock_Report").Cells(34 + i, 3)
        current_price_stock1 = Worksheets("Stock_Report").Cells(35 + i, 2)
        yesterday_price_stock1 = Worksheets("Stock_Report").Cells(34 + i, 2)
        
        Worksheets("Stock_Report").Cells(35 + i, 3).Value = yesterday_value_stock1 * current_price_stock1 / yesterday_price_stock1
        
        Worksheets("Control_Panel").range("E2").Value = Round(Worksheets("Stock_Report").range("C1290"), 2)
        
    Next i

    
    ' 6.1. Create an empty chart (stock2) only if stock 2 exists
    
    If stock2 <> "" Then
        
        date_row_length_stock2 = Worksheets(stock2).Cells(Rows.Count, 1).End(xlUp).Row
    
        Set stock2_chart = Worksheets("Stock_Report").ChartObjects.Add( _
            Top:=Worksheets("Stock_Report").Cells(1, 11).Top, _
            Left:=Worksheets("Stock_Report").Cells(1, 11).Left, _
            Width:=600, Height:=500)
    
        With stock2_chart.Chart
            .SetSourceData Source:=Worksheets(stock2).range("A1:A" & date_row_length_stock2 & " ,E1:E" & date_row_length_stock2)
            .ChartType = xlLine
            .HasTitle = True
            .ChartTitle.Text = "Price Trend for " & stock2
            .Axes(xlValue).HasMajorGridlines = False    ' Remove gridlines
        End With
        
    ' 6.2.  Add the data for stock 2 below the chart for stock 2
    ' ----Date Column: Stock 2----
    With Worksheets("Stock_Report").range("K35")
        .Value = "Date"
        .Font.Bold = True
    End With
    
    Worksheets(stock2).range("A2:A" & date_row_length_stock1).Copy Worksheets("Stock_Report").range("K36") ' Copy and paste "Date" column
    
    ' ----Closing Price Column: Stock 2----
    With Worksheets("Stock_Report").range("L35")
        .Value = stock2 & " (Closing Price)"
        .Font.Bold = True
    End With
    
    Worksheets(stock2).range("E2:E" & date_row_length_stock1).Copy Worksheets("Stock_Report").range("L36") ' Copy and paste "Close" column
    
    ' ----Investment Value Column: Stock 2----
    Dim initial_value_stock2 As Double
    Dim yesterday_value_stock2 As Double
    Dim current_price_stock2 As Double
    Dim yesterday_price_stock2 As Double
    
    initial_value_stock2 = Worksheets("Control_Panel").range("D3").Value
    
    With Worksheets("Stock_Report").range("M35")
        .Value = "Investment Value "
        .Font.Bold = True
    End With
    
    Worksheets("Stock_Report").range("M36").Value = initial_value_stock2
    
    For i = 2 To date_row_length_stock1
    
        yesterday_value_stock2 = Worksheets("Stock_Report").Cells(34 + i, 13)
        current_price_stock2 = Worksheets("Stock_Report").Cells(35 + i, 12)
        yesterday_price_stock2 = Worksheets("Stock_Report").Cells(34 + i, 12)
        
        Worksheets("Stock_Report").Cells(35 + i, 13).Value = yesterday_value_stock2 * current_price_stock2 / yesterday_price_stock2
        
        Worksheets("Control_Panel").range("E3").Value = Round(Worksheets("Stock_Report").range("M1290"), 2)
        
        ' ----Total Investment Value Column (Stock 1 + Stock 2)----
        With Worksheets("Stock_Report").range("O35")
            .Value = "Total Investment Value (Stock 1 + Stock 2)"
            .Font.Bold = True
        End With
        
        Worksheets("Stock_Report").Cells(34 + i, 15).Value = Worksheets("Stock_Report").Cells(34 + i, 13).Value + Worksheets("Stock_Report").Cells(34 + i, 3).Value
        
    Next i
        
        Worksheets("Control_Panel").range("E4").Value = Round(Worksheets("Stock_Report").range("O1290"), 2)
        
        Call GenerateCombinedChart.GenerateCombinedStockCharts
    
    Else
        Call SavingsPerformanceChart.SavingsChart_Single
        
    End If
    
    Call CalculateBasicStats.CalculateBasicStats
    
End Sub
