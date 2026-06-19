Attribute VB_Name = "DurationAndConvexity"
Function MacDur(CFRange As range, interestRate As Double) As Double
    'Calculate Macaulay Duration given a two-column table
    'where the first column represents the value of a cash flow
    'and the second column the frequency of that cash flow
    'similar to the CF worksheet in the TI BA-II Plus.
    'The first cash flow occurs at time 0
    'The interest rate is i (in decimal, e.g. 0.05 instead of 5%
    'For any cash flow value, if given no corresponding frequency, the frequency is set to be 1


    'Input Validation
    'Ensure the input range has exactly two columns
    If CFRange.Columns.Count <> 2 Then
        'Return a #VALUE! error if the range is not 2 columns
        MacDur = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Declare and initialize variables
    'Number of Rows with data
    Dim nRows As Long: nRows = Application.WorksheetFunction.CountA(CFRange.Columns(1))
    If nRows = 0 Then Exit Function
    'Numerator and Denominator of D_Mac
    Dim numerator As Double: numerator = 0
    Dim denominator As Double: denominator = 0
    'Loop counters, cash flow value, frequency and timing
    Dim i As Long, j As Long, CF As Double, freq As Long, t As Long: t = -1
    
    'Process each row in CFRange
    For i = 1 To nRows
        'Read cash flow value and frequency
        CF = CFRange.Cells(i, 1).Value
        If IsEmpty(CFRange.Cells(i, 2).Value) Then
            freq = 1 'Assume frequency is 1 if empty
        Else
            freq = CFRange.Cells(i, 2).Value
        End If
        
        'Data validation
        If Not IsNumeric(CF) Or Not IsNumeric(freq) Or freq <= 0 Then
            'Return #VALUE! if data is not numeric
            Exit Function
        End If
        
        'Add to numerator and denominator
        For j = 1 To freq
            numerator = numerator + (t + j) * CF * (1 + interestRate) ^ (-(t + j))
            denominator = denominator + CF * (1 + interestRate) ^ (-(t + j))
            'MsgBox numerator & "; " & denominator
        Next j
        
        t = t + freq
    Next i
    
    If denominator = 0 Then
        MacDur = CVErr(xlErrDiv)
    Else
        MacDur = numerator / denominator
    End If
    

End Function



Function ModDur(CFRange As range, interestRate As Double) As Double
    'Calculate Modified Duration given a two-column table
    'where the first column represents the value of a cash flow
    'and the second column the frequency of that cash flow
    'similar to the CF worksheet in the TI BA-II Plus.
    'The first cash flow occurs at time 0
    'The interest rate is i (in decimal, e.g. 0.05 instead of 5%
    
    
    ModDur = MacDur(CFRange, interestRate) / (1 + interestRate)
    
    
End Function



Function MacConvexity(CFRange As range, interestRate As Double) As Double
    'Calculate Macaulay Convexity given a two-column table
    'where the first column represents the value of a cash flow
    'and the second column the frequency of that cash flow
    'similar to the CF worksheet in the TI BA-II Plus.
    'The first cash flow occurs at time 0
    'The interest rate is i (in decimal, e.g. 0.05 instead of 5%
    'For any cash flow value, if given no corresponding frequency, the frequency is set to be 1


    'Input Validation
    'Ensure the input range has exactly two columns
    If CFRange.Columns.Count <> 2 Then
        'Return a #VALUE! error if the range is not 2 columns
        MacConvexity = CVErr(xlErrValue)
        Exit Function
    End If
    
    'Declare and initialize variables
    'Number of Rows with data
    Dim nRows As Long: nRows = Application.WorksheetFunction.CountA(CFRange.Columns(1))
    If nRows = 0 Then Exit Function
    'Numerator and Denominator of D_Mac
    Dim numerator As Double: numerator = 0
    Dim denominator As Double: denominator = 0
    'Loop counters, cash flow value, frequency and timing
    Dim i As Long, j As Long, CF As Double, freq As Long, t As Long: t = -1
        
    'Process each row in CFRange
    For i = 1 To nRows
        'Read cash flow value and frequency
        CF = CFRange.Cells(i, 1).Value
        If IsEmpty(CFRange.Cells(i, 2).Value) Then
            freq = 1 'Assume frequency is 1 if empty
        Else
            freq = CFRange.Cells(i, 2).Value
        End If
        
        'Data validation
        If Not IsNumeric(CF) Or Not IsNumeric(freq) Or freq <= 0 Then
            'Return #VALUE! if data is not numeric
            MacConvexity = CVErr(xlErrValue)
            Exit Function
        End If
        
        'Add to numerator and denominator
        For j = 1 To freq
            numerator = numerator + (t + j) ^ 2 * CF * (1 + interestRate) ^ (-(t + j))
            denominator = denominator + CF * (1 + interestRate) ^ (-(t + j))
            'MsgBox numerator & "; " & denominator
        Next j
        
        t = t + freq
    Next i
    
    If denominator = 0 Then
        MacConvexity = CVErr(xlErrDiv)
    Else
        MacConvexity = numerator / denominator
    End If
    

End Function



Function ModConvexity(CFRange As range, interestRate As Double) As Double
    'Calculate Modified Convexity given a two-column table
    'where the first column represents the value of a cash flow
    'and the second column the frequency of that cash flow
    'similar to the CF worksheet in the TI BA-II Plus.
    'The first cash flow occurs at time 0
    'The interest rate is i (in decimal, e.g. 0.05 instead of 5%
    
    
    ModConvexity = (MacDur(CFRange, interestRate) + MacConvexity(CFRange, interestRate)) _
                    / (1 + interestRate) ^ 2
    
    
End Function
