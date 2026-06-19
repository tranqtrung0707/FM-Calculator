Attribute VB_Name = "CreateAmortTable"
Sub CreateAmortTable()

    
    'Declare variables
    Dim loanAmount As Double
    Dim interestRate As Double
    Dim numPayments As Long
    Dim payment As Double
    Dim balance As Double
    Dim interestPayment As Double
    Dim principalPayment As Double
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Long

    'Assign user inputs to variables
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Amortization")
    numPayments = ws.range("L_N_val").Value
    interestRate = ws.range("L_I_val").Value / 100 'convert to decimal
    loanAmount = ws.range("L_L_val").Value
    payment = ws.range("L_Pmt_val").Value
    
    'Check for valid inputs
    If numPayments <= 0 Or interestRate < 0 Then
        MsgBox "Number of payments must be positive and interest rate non-negative.", vbCritical
        Exit Sub
    End If

    'Clear range for data
    ws.range("B3:G1048576").ClearContents
    ws.range("RemainingBal").ClearContents
    
    'Clear border
    ws.range("B2:G1048576").Borders.LineStyle = xlNone
    
    'Set the first row to write data, which is Row 3
    startRow = 3
    
    'The first beginning balance is not calculated like the rest
    ws.Cells(startRow, 3).Formula = "= L_L_val"

    'Main loop
    For i = 0 To numPayments - 1
        ws.Cells(startRow + i, 2).Value = i + 1
        If i > 0 Then ws.Cells(startRow + i, 3).FormulaR1C1 = "= R[-1]C[4]"
        ws.Cells(startRow + i, 4).Formula = "= L_Pmt_val"
        ws.Cells(startRow + i, 5).FormulaR1C1 = "= RC[-2] * L_I_val/100"
        'MsgBox "pmt=" & payment & ", balance=" & balance & ", I=" & interestRate
        ws.Cells(startRow + i, 6).FormulaR1C1 = "= RC[-2] - RC[-1]"
        ws.Cells(startRow + i, 7).FormulaR1C1 = "= RC[-4] - RC[-1]"
    Next i
    
    'Write down the remaining balance
    ws.range("RemainingBal").Formula = "=G" & (startRow + numPayments - 1)
    
    'Draw a border around the entire table
    With ws.range(ws.Cells(startRow - 1, 2), ws.Cells(startRow + numPayments - 1, 7))
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    End With
    
    'Autofit data columns
    ws.range("B:G").Columns.AutoFit
    
    MsgBox "Amortization table generated successfully!", vbInformation


End Sub
