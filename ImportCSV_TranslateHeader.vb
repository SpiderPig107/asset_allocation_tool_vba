Attribute VB_Name = "ImportCSV_TranslateHeader"
Sub ImportAllCSV()

Dim folderpath As String
Dim filename As String
Dim tempworkbook As Workbook
Dim masterworkbook As Workbook
Dim sheetname As String

' 1. Set the folder path
folderpath = "/Users/pc1/Documents/01_TISE/03_PUEB_Poland/04_Data Analysis with VBA/Project/Data/"

' 2. Find the first CSV in the folder
filename = Dir(folderpath & "*.csv")
' Dir function goes to that folder, finds the file, and sets the variable fileName to exactly "aapl_us_d.csv". It strips away the folder path.

Set masterworkbook = ThisWorkbook

' 3. Loop through all files until no more is found
Do While filename <> ""
    ' 3.1. Open the CSV file
    Set tempworkbook = Workbooks.Open(folderpath & filename)
    
    ' 3.2. Create a new sheet (without .csv in the sheet name)
    sheetname = Left(filename, InStr(filename, "_") - 1)
    ' Left(filename, ...) This tells VBA to start at the beginning of the name and grab only that many characters.
    ' InStr is In-String Search. It finds the position of the dot. In aapl.csv, the dot is at position 5.
    
    ' 3.3. Copy the sheet from tempworkbook to masterworkbook
    tempworkbook.Sheets(1).Copy After:=masterworkbook.Sheets(masterworkbook.Sheets.Count)
    
    ' .Copy duplicates the sheet
    ' After:= is a Named Argument. It is a specific instruction telling Excel exactly where to drop the new sheet
    ' The Dot (.): This is for ownership. importWorkbook.Sheets(1) means "The first sheet belonging to that workbook."
    ' The Space: This tells VBA, "I am done naming the object and the action; now I am going to give you the specific details on how to do it."
    ' If you wrote .Copy.After, VBA would think After is a property inside the Copy command, which doesn't exist.
    ' Instead, Copy is the command, and After is a "setting" for that command.
    
    ' 3.4. Rename the new sheetname in masterworkbook (using sheetname in 3.2)
    masterworkbook.Sheets(masterworkbook.Sheets.Count).Name = sheetname
    
    ' 3.5. Close the tempworkbook
    tempworkbook.Close SaveChanges:=False
    
    ' 3.6. Find the next file in the folderpath
    filename = Dir()
    ' Because you didn't provide a path this time, VBA knows you mean: "Go back to that same folder from before and give me the next file name in the list."

Loop

MsgBox "Success: Import CSV files completed."

End Sub

Sub TranslateHeader()

Dim sheet As Worksheet

    ' Loop through every sheet in the masterworkbook
    For Each sheet In ThisWorkbook.Worksheets
        
        ' Skip the "Control_Panel" sheet
        If sheet.Name <> "Control_Panel" Then
        
            ' Replace Polish names with English names in Row 1
            sheet.range("A1").Value = "Date"            ' Data
            sheet.range("B1").Value = "Open"           ' Otwarcie
            sheet.range("C1").Value = "High"            ' Najwyzscy
            sheet.range("D1").Value = "Low"             ' Najnizszy
            sheet.range("E1").Value = "Close"           ' Zamkniecie
            sheet.range("F1").Value = "Volume"        ' Wolumen
            
        End If
    Next sheet

MsgBox "Translation completed! Header is now in English!"

End Sub
