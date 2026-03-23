Attribute VB_Name = "InsertCompanyName1_2"
Option Explicit

Sub InsertCompanyName1()
    
    Dim stock1_input As String
    Dim ws_control_panel As Worksheet
    Dim menu_list1 As String
    Dim code As String
    Dim stock1_weight
    
    Set ws_control_panel = Worksheets("Control_Panel")
    
    ' --- STARTING STEP ---
    ' Clear A2:A3 (Names), B2:B3 (Codes), and C2:C3 (Weights)
    ws_control_panel.range("A2:E3").ClearContents
    ws_control_panel.range("E4").ClearContents
    
    ' Clear Old Charts
    Dim char_obj As ChartObject
    
    If ws_control_panel.ChartObjects.Count > 0 Then
        For Each char_obj In ws_control_panel.ChartObjects
            char_obj.Delete
        Next char_obj
    End If
    
    ' 1. Build the list of companies to show in the prompt
    menu_list1 = "Choose two stocks." & vbCrLf & _
                        vbCrLf & _
                        "It must be one of these companies:" & vbCrLf & _
                        "Amazon" & vbCrLf & _
                        "Apple" & vbCrLf & _
                        "Boeing" & vbCrLf & _
                        "Disney" & vbCrLf & _
                        "IBM" & vbCrLf & _
                        "McDonald's" & vbCrLf & _
                        "Merck" & vbCrLf & _
                        "Microsoft" & vbCrLf & _
                        "Nike" & vbCrLf & _
                        "Tesla" & vbCrLf & _
                        "Visa" & vbCrLf & _
                        "Walmart" & vbCrLf & _
                        vbCrLf & _
                        "Enter your Stock 1"
    
    ' --- STOCK 1: COMPANY NAME ---
    ' 2. Ask for the first company name
    Do
        stock1_input = InputBox(menu_list1, "Enter Name of Stock 1 (Step 1 of 4)")
        If stock1_input = "" Then Exit Sub ' Exit if user cancels
    
    ' 2. Create the logic to convert name to company code (Stock1)
        Select Case LCase(Trim(stock1_input))
            Case "amazon":       code = "amzn"
            Case "apple":           code = "aapl"
            Case "boeing":         code = "ba"
            Case "disney":         code = "dis"
            Case "ibm":             code = "ibm"
            Case "mcdonalds", "mcdonald's":     code = "mcd"
            Case "merck":         code = "mrk"
            Case "microsoft":   code = "msft"
            Case "nike":           code = "nke"
            Case "tesla":          code = "tsla"
            Case "visa":           code = "v"
            Case "walmart":    code = "wmt"
            Case Else:            code = "unknown"
        End Select
    
    ' 3. Validate user input for stock 1
        If code <> "unknown" Then
            ws_control_panel.range("A2").Value = stock1_input
            ws_control_panel.range("B2").Value = code
        Else
            MsgBox "Please enter name exactly as it appears on the menu for Stock 1", vbExclamation, "Invalid Weight"
        End If
    
    Loop While code = "unknown"
    
    MsgBox "You have selected: " & stock1_input, vbInformation, "Stock 1 Name: Success"
    

    ' --- STOCK 1: WEIGHT ---
    ' 4. Ask for the weight of Stock 1
    Do
        stock1_weight = InputBox("Enter the weight for " & stock1_input & " (0 to 100):", "Enter Weight of Stock 1 (Step 2 of 4)")
        If stock1_weight = "" Then Exit Sub ' Exit if user cancels
        
        If IsNumeric(stock1_weight) Then
            stock1_weight = CDbl(stock1_weight) ' Convert to Double
            
            If stock1_weight >= 0 And stock1_weight <= 100 Then
                ws_control_panel.range("C2").Value = stock1_weight / 100
                ws_control_panel.range("D2").Value = ws_control_panel.range("B10") * stock1_weight / 100
                Exit Do ' Success exit
            End If
        
        End If
        
        MsgBox "Enter a number between 0 and 100.", vbExclamation, "Invalid Weight"
    Loop ' Loop back to InputBox of stock1_weight
    
    MsgBox "You have selected " & stock1_input & " with " & stock1_weight & "% allocation.", vbInformation, "Stock 1 Name & Weight: Success"
    
    ' --- DECISION GATE ---
    ' If weight is 100, we don't need Stock 2. Skip to Charts.
    
    If stock1_weight = 100 Then
        MsgBox "Portfolio is 100% " & stock1_input & ". Skipping Stock 2.", vbInformation, "Stock 2 not enabled"
        ws_control_panel.range("D2").Value = ws_control_panel.range("B10") * stock1_weight / 100
        Call GenerateIndividualStockCharts.GenerateStockCharts
    Else
        ' Otherwise, proceed to ask for the second stock as usual
        Call InsertCompanyName2
    End If
    
End Sub


Sub InsertCompanyName2()

Dim stock2_input As String
Dim ws_control_panel As Worksheet
Dim menu_list2 As String
Dim code As String
Dim stock2_weight As Variant
Dim weight1_actual As Double
Dim max_allowed As Double
    
Set ws_control_panel = Worksheets("Control_Panel")

' 1. Build the list of companies to show in the prompt
menu_list2 = "Choose your second stock." & vbCrLf & _
                        vbCrLf & _
                        "It must be one of these companies:" & vbCrLf & _
                        "Amazon" & vbCrLf & _
                        "Apple" & vbCrLf & _
                        "Boeing" & vbCrLf & _
                        "Disney" & vbCrLf & _
                        "IBM" & vbCrLf & _
                        "McDonald's" & vbCrLf & _
                        "Merck" & vbCrLf & _
                        "Microsoft" & vbCrLf & _
                        "Nike" & vbCrLf & _
                        "Tesla" & vbCrLf & _
                        "Visa" & vbCrLf & _
                        "Walmart" & vbCrLf & _
                        vbCrLf & _
                        "Enter your Stock 2"

' --- STOCK 2: COMPANY NAME ---
' 2. Ask for the second company name
Do
    stock2_input = InputBox(menu_list2, "Enter Name of Stock 2 (Step 3 of 4)")
    If stock2_input = "" Then Exit Sub ' Exit if user cancels
    
    ' 2. Create the logic to convert name to company code (Stock1)
    Select Case LCase(Trim(stock2_input))
        Case "amazon":       code = "amzn"
        Case "apple":           code = "aapl"
        Case "boeing":         code = "ba"
        Case "disney":         code = "dis"
        Case "ibm":             code = "ibm"
        Case "mcdonalds", "mcdonald's":     code = "mcd"
        Case "merck":         code = "mrk"
        Case "microsoft":   code = "msft"
        Case "nike":           code = "nke"
        Case "tesla":          code = "tsla"
        Case "visa":           code = "v"
        Case "walmart":    code = "wmt"
        Case Else:            code = "unknown"
    End Select
    
    ' 3. Validate user input for stock 2
    If code <> "unknown" Then
        ws_control_panel.range("A3").Value = stock2_input
        ws_control_panel.range("B3").Value = code
    Else
        MsgBox "Please enter name exactly as it appears on the menu for Stock 2", vbExclamation, "Invalid Name"
    End If
    
Loop While code = "unknown"

MsgBox "You have selected: " & stock2_input, vbInformation, "Stock 2 Name: Success"
    
' --- STOCK 2: WEIGHT ---
weight1_actual = ws_control_panel.range("C2").Value * 100
stock2_weight = 100 - weight1_actual
ws_control_panel.range("C3").Value = stock2_weight / 100
ws_control_panel.range("D3").Value = ws_control_panel.range("B10") * stock2_weight / 100
    
MsgBox "You have selected " & stock2_input & " with " & stock2_weight & "% allocation." & vbCrLf & _
            "Please wait while it calculates.", vbInformation, "Stock 2 Name & Weight: Success"
        
Call GenerateIndividualStockCharts.GenerateStockCharts
        
End Sub
        
        
