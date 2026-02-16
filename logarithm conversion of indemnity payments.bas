Attribute VB_Name = "Module1"
Option Explicit

Public Sub CreateLogIndemnity_Clean()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim indemnityValue As Double
    
    ' Reference Clean sheet directly
    Set ws = ThisWorkbook.Worksheets("Clean")
    
    ' Find last used row in Column D
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    ' Create header in Column G
    ws.Range("G1").Value = "Log_Indemnity"
    
    ' Loop through indemnity values
    For i = 2 To lastRow
        
        indemnityValue = ws.Cells(i, "D").Value
        
        If indemnityValue > 0 Then
            ws.Cells(i, "G").Value = WorksheetFunction.Log10(indemnityValue)
        Else
            ws.Cells(i, "G").Value = ""
        End If
        
    Next i
    
    MsgBox "Log transformation complete in Clean sheet.", vbInformation

End Sub

