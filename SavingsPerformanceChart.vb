Attribute VB_Name = "SavingsPerformanceChart"
Option Explicit

Sub SavingsChart_Both() ' This is with Stock 1 and 2

    ' 1. Create an empty chart
    Dim savings_chart_both As ChartObject
    Set savings_chart_both = Worksheets("Control_Panel").ChartObjects.Add( _
        Top:=Worksheets("Control_Panel").Cells(9, 10).Top, _
        Left:=Worksheets("Stock_Report").Cells(9, 10).Left, _
        Width:=600, Height:=500)
        
    ' 2. Measure the row length of dates from the individual stock sheet
    Dim stock1 As String, stock2 As String
    Dim weight1 As Double, weight2 As Double
    Dim date_row_length_stock1 As Long
    
    stock1 = Worksheets("Control_Panel").range("B2").Value
    stock2 = Worksheets("Control_Panel").range("B3").Value
    
    weight1 = Worksheets("Control_Panel").range("C2").Value * 100
    weight2 = Worksheets("Control_Panel").range("C3").Value * 100
    
    date_row_length_stock1 = Worksheets(stock1).Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 3. Link the data range to the empty chart (stock1 & 2)
    With savings_chart_both.Chart
        .SetSourceData Source:=Worksheets("Stock_Report").range("A35:A" & date_row_length_stock1 & " ,O35:O" & date_row_length_stock1)
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Total Investment Value          " & "(" & stock1 & " at " & weight1 & "% " & " and " & stock2 & " at " & weight2 & "%" & ")"
        .Axes(xlValue).HasMajorGridlines = False    ' Remove gridlines
    End With

End Sub


Sub SavingsChart_Single() ' This is with Stock 1 only

    ' 1. Create an empty chart
    Dim savings_chart_single As ChartObject
    Set savings_chart_single = Worksheets("Control_Panel").ChartObjects.Add( _
        Top:=Worksheets("Control_Panel").Cells(9, 10).Top, _
        Left:=Worksheets("Stock_Report").Cells(9, 10).Left, _
        Width:=600, Height:=500)
        
    ' 2. Measure the row length of dates from the individual stock sheet
    Dim stock1 As String
    Dim weight1 As Double
    Dim date_row_length_stock1 As Long
    
    stock1 = Worksheets("Control_Panel").range("B2").Value
    
    weight1 = Worksheets("Control_Panel").range("C2").Value * 100
    
    date_row_length_stock1 = Worksheets(stock1).Cells(Rows.Count, 1).End(xlUp).Row
    
    ' 3. Link the data range to the empty chart (stock1 & 2)
    With savings_chart_single.Chart
        .SetSourceData Source:=Worksheets("Stock_Report").range("A35:A" & date_row_length_stock1 & " ,C35:C" & date_row_length_stock1)
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Total Investment Value          " & "(" & stock1 & " at " & weight1 & "% " & ")"
        .Axes(xlValue).HasMajorGridlines = False    ' Remove gridlines
    End With

End Sub
