Attribute VB_Name = "Module1"
Option Explicit

Public Sub BuildClean_FixedCols_V2()

    Dim wsR As Worksheet, wsC As Worksheet
    Dim lastRow As Long, r As Long

    Set wsR = ThisWorkbook.Worksheets("Raw")
    Set wsC = ThisWorkbook.Worksheets("Clean")

    ' last row based on Claim Number column (B)
    lastRow = wsR.Cells(wsR.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Raw has no data in column B (Claim Number).", vbCritical
        Exit Sub
    End If

    ' clear only Clean
    wsC.Cells.Clear

    ' headers
    wsC.Range("A1").Value = "Claim id"
    wsC.Range("B1").Value = "Final disposition date"
    wsC.Range("C1").Value = "Severity text"
    wsC.Range("D1").Value = "Indemnity Paid"
    wsC.Range("E1").Value = "Post Reform"
    wsC.Range("F1").Value = "Severity Code"

    For r = 2 To lastRow
        wsC.Cells(r, "A").Value = wsR.Cells(r, "B").Value          ' Claim Number
        wsC.Cells(r, "B").Value = NormalizeDate(wsR.Cells(r, "AA").Value) ' Final Disposition
        wsC.Cells(r, "C").Value = wsR.Cells(r, "O").Value          ' Severity
        wsC.Cells(r, "D").Value = CleanMoney(wsR.Cells(r, "X").Value)    ' Indemnity

        wsC.Cells(r, "E").Value = PostFlag(wsC.Cells(r, "B").Value)
        wsC.Cells(r, "F").Value = SeverityCode(wsC.Cells(r, "C").Value)
    Next r

    wsC.Columns("D").NumberFormat = "#,##0"
    wsC.Columns("B").NumberFormat = "m/d/yyyy"
    wsC.Columns("A:F").AutoFit

    MsgBox "Done. Processed " & (lastRow - 1) & " rows.", vbInformation

End Sub

Private Function CleanMoney(v As Variant) As Variant
    Dim s As String
    s = CStr(v)
    s = Replace(s, "$", "")
    s = Replace(s, ",", "")
    s = Trim$(s)

    If s = "" Then
        CleanMoney = ""
    ElseIf IsNumeric(s) Then
        CleanMoney = CDbl(s)
    Else
        CleanMoney = v
    End If
End Function

Private Function NormalizeDate(v As Variant) As Variant
    On Error GoTo Fail
    If IsDate(v) Then NormalizeDate = CDate(v): Exit Function
Fail:
    NormalizeDate = v
End Function

Private Function PostFlag(dateVal As Variant) As Variant
    On Error GoTo Fail
    If CDate(dateVal) >= DateSerial(2023, 3, 24) Then
        PostFlag = 1
    Else
        PostFlag = 0
    End If
    Exit Function
Fail:
    PostFlag = ""
End Function

Private Function SeverityCode(sevText As Variant) As Variant
    Dim s As String
    s = LCase$(CStr(sevText))

    If InStr(s, "emotional") > 0 Then SeverityCode = 1: Exit Function
    If InStr(s, "temporary: slight") > 0 Then SeverityCode = 2: Exit Function
    If InStr(s, "temporary: minor") > 0 Then SeverityCode = 3: Exit Function
    If InStr(s, "temporary: major") > 0 Then SeverityCode = 4: Exit Function
    If InStr(s, "permanent: minor") > 0 Then SeverityCode = 5: Exit Function
    If InStr(s, "permanent: significant") > 0 Then SeverityCode = 6: Exit Function
    If InStr(s, "permanent: major") > 0 Then SeverityCode = 7: Exit Function
    If InStr(s, "permanent: grave") > 0 Then SeverityCode = 8: Exit Function
    If InStr(s, "permanent: death") > 0 Then SeverityCode = 9: Exit Function

    SeverityCode = ""
End Function

