Attribute VB_Name = "Function_SemiStd"
Option Explicit

' Semi Variance = sum of ((value below mean - mean)^2) / number of observations below mean
'' Semi standard deviation = square root(semi variance)

Function semi_std(selected_range As range) As Double
    
    Dim sum_sqr_neg_dev As Double
    Dim count_below_mean As Long
    Dim mean As Double
    Dim cell As range
    Dim semi_var As Double
    
    ' 1. Compute Semi-Variance
    ' 1.1. Compute SUM(value below mean - mean)^2; and number observations below mean
    sum_sqr_neg_dev = 0 ' sum of squared negative deviations (value below mean - mean)^2
    count_below_mean = 0
    mean = Application.WorksheetFunction.Average(selected_range)
    
    For Each cell In selected_range
    
        If cell.Value < mean Then
            sum_sqr_neg_dev = sum_sqr_neg_dev + (cell.Value - mean) ^ 2
            count_below_mean = count_below_mean + 1
        End If
        
    Next cell
    
    ' 1.2. Compute Semi-Variance
    If count_below_mean > 0 Then
        semi_var = sum_sqr_neg_dev / count_below_mean
    Else
        semi_var = 0    ' If there are no values below mean, then semi_var is 0
    End If
    
    ' 2. Compute Semi-Standard Deviation
    semi_std = Sqr(semi_var) ' VBA uses Sqr() for square root
    
End Function
