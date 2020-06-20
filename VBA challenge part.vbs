Sub tickertotaler_moderate():

For Each ws In ThisWorkbook.Worksheets
    ws.Cells(6, 14).Value = "Greatest % increase"
    ws.Cells(7, 14).Value = "Greatest % decrease"
    ws.Cells(8, 14).Value = "Greatest total volume"
    ws.Cells(6, 15).Value = WorksheetFunction.Max(ws.Range("K2:K3005"))
    ws.Cells(7, 15).Value = WorksheetFunction.Min(ws.Range("K2:K3005"))
    ws.Cells(8, 15).Value = WorksheetFunction.Max(ws.Range("L2:L3005"))
    Next
    
End Sub
