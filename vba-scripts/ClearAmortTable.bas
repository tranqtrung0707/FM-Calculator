Attribute VB_Name = "ClearAmortTable"
Sub ClearAmortTable()


    Set ws = ThisWorkbook.Sheets("Amortization")

    'Clear data cells
    ws.range("B3:G1048576").ClearContents
    ws.range("RemainingBal").ClearContents
    
    'Clear border
    ws.range("B2:G1048576").Borders.LineStyle = xlNone


End Sub
