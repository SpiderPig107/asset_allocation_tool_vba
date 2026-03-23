Attribute VB_Name = "ExportCode"
Option Explicit

Sub SaveCodeModules()
' This code Exports all VBA modules

Dim i As Integer
Dim sName As String
Dim exportFolder As String

' Define your specific GitHub path here - Ensure it ends with a slash /
exportFolder = "/Users/pc1/Documents/GitHub/asset_allocation_tool_vba/"
    
    With ThisWorkbook.VBProject
    
        For i = 1 To .VBComponents.Count
             sName = .VBComponents(i).CodeModule.Name
             .VBComponents(i).Export exportFolder & sName & ".vba"
        Next i
        
    End With

End Sub
