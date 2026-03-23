Attribute VB_Name = "CalculateBasicStats"
Option Explicit

Sub CalculateBasicStats()

Dim ws_control_panel As Worksheet

Dim stock1 As String, stock2 As String

Dim date_row_length_stock1 As Long, date_row_length_stock2 As Long

Dim price_range_stock1 As range, price_range_stock2 As range

Dim mean_stock1 As Double, mean_stock2 As Double

Dim std_stock1 As Double, std_stock2 As Double

Dim semi_std_stock1 As Double, semi_std_stock2 As Double


'------------STOCK 1---------------------
    ' 1. Identify the two company from Control Panel
    Set ws_control_panel = ThisWorkbook.Sheets("Control_Panel")
    stock1 = ws_control_panel.range("B2").Value
    
    ' 2. Measure the row length of dates from the individual stock sheet
    date_row_length_stock1 = Worksheets(stock1).Cells(Rows.Count, 1).End(xlUp).Row

    ' 3. Locate the data range for "Close"
    Set price_range_stock1 = Worksheets(stock1).range("E2:E" & date_row_length_stock1)
    
    ' 4. Calculate the Mean, Std and Semi_Std for Stock 1
    mean_stock1 = Application.WorksheetFunction.Average(price_range_stock1)
    std_stock1 = Application.WorksheetFunction.StDev(price_range_stock1)
    semi_std_stock1 = semi_std(price_range_stock1)
    
    ' 6. Add the text box containing the stats in the chart
    
    ' 6.1. Add the text box to the first Individual Chart (Stock1)
    With Worksheets("Stock_Report").ChartObjects(1).Chart.Shapes.AddTextbox(1, 35, 5, 110, 25)
    
    ' AddTextBox(Orientation, Left, Top, Width, Height).
    ' 1 (Orientation) - It means your text will read normally from left to right.
    ' 10 (Left) - This is the horizontal "X" coordinate. It places the box 10 points away from the left edge of the chart.
    ' 10 (Top) - This is the vertical "Y" coordinate. It places the box 10 points down from the top edge of the chart.
    ' 110 (Width) - This is the size of the box from left to right (110 points wide).
    ' 25 (Height) - This is the size of the box from top to bottom (25 points tall).
        
        .TextFrame.Characters.Text = _
            "Mean: " & Format(mean_stock1, "0.00") & vbCrLf & _
            "Std: " & Format(std_stock1, "0.00") & vbCrLf & _
            "Semi-Std: " & Format(semi_std_stock1, "0.00")
                
                With .TextFrame.Characters.Font
                    .Size = 11
                    .Color = RGB(255, 0, 0)         ' Set font color to Red
                    .Bold = True
                End With
    
        .TextFrame.AutoSize = True      ' Ensure the text frame doesn't hide
    
    End With
    
    ' 6.3. Add the text box to the Combined Chart (Stock1)
    
    ' 6.3.1. Stock 1 (Left side of chart)
If Worksheets("Control_Panel").range("B3") <> "" Then

    With Worksheets("Combined_Stock_Report").ChartObjects(1).Chart.Shapes.AddTextbox(1, 35, 5, 110, 25)
        
        .TextFrame.Characters.Text = _
            stock1 & vbCrLf & _
            "Mean: " & Format(mean_stock1, "0.00") & vbCrLf & _
            "Std: " & Format(std_stock1, "0.00") & vbCrLf & _
            "Semi-Std: " & Format(semi_std_stock1, "0.00")
            
                With .TextFrame.Characters.Font
                    .Size = 11
                    .Color = RGB(255, 0, 0)         ' Set font color to Red
                    .Bold = True
                End With
    
        .TextFrame.AutoSize = True      ' Ensure the text frame doesn't hide
    
    End With

End If

'----------------------STOCK 2---------------------

If ws_control_panel.range("B3").Value <> "" Then

    ' 1. Identify the second company from Control Panel
    stock2 = ws_control_panel.range("B3").Value
    
    ' 2. Measure the row length of dates from the individual stock sheet
    date_row_length_stock2 = Worksheets(stock2).Cells(Rows.Count, 1).End(xlUp).Row

    ' 3. Locate the data range for "Close"
    Set price_range_stock2 = Worksheets(stock2).range("E2:E" & date_row_length_stock2)
    
    ' 4. Calculate the Mean, Std and Semi_Std for Stock 2
    mean_stock2 = Application.WorksheetFunction.Average(price_range_stock2)
    std_stock2 = Application.WorksheetFunction.StDev(price_range_stock2)
    semi_std_stock2 = semi_std(price_range_stock2)
    
    ' 6. Add the text box containing the stats in the chart
    
    ' 6.1. Add the text box to the first Individual Chart (Stock2)
    With Worksheets("Stock_Report").ChartObjects(2).Chart.Shapes.AddTextbox(1, 35, 5, 110, 25)
        
        .TextFrame.Characters.Text = _
            "Mean: " & Format(mean_stock2, "0.00") & vbCrLf & _
            "Std: " & Format(std_stock2, "0.00") & vbCrLf & _
            "Semi-Std: " & Format(semi_std_stock2, "0.00")
                
                With .TextFrame.Characters.Font
                    .Size = 11
                    .Color = RGB(255, 0, 0)         ' Set font color to Red
                    .Bold = True
                End With
    
        .TextFrame.AutoSize = True      ' Ensure the text frame doesn't hide
    
    End With

    ' 6.3. Add the text box to the Combined Chart (Stock2)
    
    ' 6.3.1. Stock 2 (Right side of chart)
    With Worksheets("Combined_Stock_Report").ChartObjects(1).Chart.Shapes.AddTextbox(1, 500, 5, 110, 25)
        
        .TextFrame.Characters.Text = _
            stock2 & vbCrLf & _
            "Mean: " & Format(mean_stock2, "0.00") & vbCrLf & _
            "Std: " & Format(std_stock2, "0.00") & vbCrLf & _
            "Semi-Std: " & Format(semi_std_stock2, "0.00")
                
                With .TextFrame.Characters.Font
                    .Size = 11
                    .Color = RGB(255, 0, 0)         ' Set font color to Brown
                    .Bold = True
                End With
    
        .TextFrame.AutoSize = True      ' Ensure the text frame doesn't hide
    
    End With
    
End If
    
End Sub
