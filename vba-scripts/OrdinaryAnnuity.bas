Attribute VB_Name = "OrdinaryAnnuity"
'Calculate time value of money variables for an ordinary annuity

Sub OrdinaryAnnuity()
    
    'Refer to the current worksheet with a shorter name
    Set ws = ThisWorkbook.Sheets("Annuities")
    
    Static LastColoredCell As range
    
    'Clearing Previous Color
    If Not LastColoredCell Is Nothing Then
        On Error Resume Next
        LastColoredCell.Interior.ColorIndex = xlNone
        On Error GoTo 0
    End If
    Dim computedCell As range

    'Declare variables for each input cell
    Dim Mode_cell As range, N_cell As range, I_cell As range, PV_cell As range, PMT_cell As range, FV_cell As range
    
    'Assign the ranges to specific cells
    Set Mode_cell = ws.range("Mode_val")
    Set N_cell = ws.range("N_val")
    Set I_cell = ws.range("I_val")
    Set PV_cell = ws.range("PV_val")
    Set PMT_cell = ws.range("PMT_val")
    Set FV_cell = ws.range("FV_val")
    
    'Variables to hold the input and computed values
    Dim Mode As Integer, N As Variant 'N can be a number or "inf": denoting perpetuity
    Dim i As Double 'Interest rate will be in decimal form (e.g., 0.05 for 5%)
    Dim PV As Double, PMT As Double, FV As Double
    
    'Counter for blank cells
    Dim blankCount As Integer
    blankCount = 0
    
    'Variable to store the address of the blank cell
    Dim blankCellAddress As String
    
    '#####################################################
    'Check each cell and count the number of blank cells, store the address of the blank cell
    If IsEmpty(N_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "N"
    ElseIf Not IsNumeric(N_cell.Value) Then
        If LCase(N_cell.Value) = "inf" Then
            N = "inf"
        Else
            MsgBox "N must be a number. Please correct your input.", vbCritical
            Exit Sub 'Stop execution if invalid input
        End If
    Else
        N = N_cell.Value
    End If
    
    If IsEmpty(I_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "I"
    ElseIf (Not IsNumeric(I_cell.Value)) Or (I_cell.Value < -100) Then
        MsgBox "I must be a number greater than or equal to -100. Please correct your input.", vbCritical
        Exit Sub
    Else
        i = I_cell.Value / 100 'Convert percentage to decimal for calculation
    End If
    
    If IsEmpty(PV_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "PV"
    ElseIf Not IsNumeric(PV_cell.Value) Then
        MsgBox "PV must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        PV = PV_cell.Value
    End If

    If IsEmpty(PMT_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "PMT"
    ElseIf Not IsNumeric(PMT_cell.Value) Then
        MsgBox "PMT must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        PMT = PMT_cell.Value
    End If
    
    If IsEmpty(FV_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "FV"
    ElseIf Not IsNumeric(FV_cell.Value) Then
        MsgBox "FV must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        FV = FV_cell.Value
    End If
    
    '####################################################
    'Input validation
    If blankCount = 0 Then
        MsgBox "Please leave exactly one cell blank to compute.", vbExclamation
        Exit Sub
    ElseIf blankCount > 1 Then
        MsgBox "Please leave only on cell blank to compute.", vbExclamation
        Exit Sub
    End If
    
    'Basic Error Handling
    On Error GoTo ErrorHandler
    
    '#######################################################
    'Perform calculation
    
    'MsgBox "N: " & N & vbCrLf & "I (decimal): " & I & vbCrLf & "PV: " & PV & vbCrLf & "PMT: " & PMT & vbCrLf & "FV: " & FV & vbCrLf & vbCrLf & "Blank Count: " & blankCount & vbCrLf & "Blank Cell to Compute: " & blankCellAddress, vbInformation, "Debugging Input Values"
    
    If Mode_cell.Value = "BGN" Then Mode = 1 Else Mode = 0
    
    'Check if N is "inf" (Perpetuity)
    If N = "inf" Then
        'FV must be zero for a perpetuity
        If FV <> 0 Then
            MsgBox "Future Value (FV) must be zero for a perpetuity", vbCritical
            FV_cell.Value = 0 'Clear FV as it should be 0
        End If
        
        'Disable FV calculation
        If blankCellAddress = "FV" Then
            MsgBox "Cannot Compute FV for a perpetuity. Please leave I, PV, or PMT blank.", vbCritical
            Exit Sub
        End If

        'Perform perpetuity calculations
        Select Case blankCellAddress
            Case "I"
                'PV and PMT must have the opposite sign and PV<>0
                If (PMT * PV > 0) Or PV = 0 Then
                    MsgBox "Cannot compute I with the given PMT and PV.", vbCritical
                    Exit Sub
                End If
                If Mode = 0 Then
                    I_cell.Value = (PMT / -PV) * 100
                Else
                    I_cell.Value = (PMT / (-PV - PMT)) * 100
                End If
                Set computedCell = I_cell
            Case "PV"
                'I must be positive
                If i <= 0 Then
                    MsgBox "Cannot compute perpetuity if interest rate (I) is not positive.", vbCritical
                    Exit Sub
                End If
                If Mode = 0 Then
                    PV_cell.Value = -PMT / i
                Else
                    PV_cell.Value = -PMT * (1 + 1 / i)
                End If
                Set computedCell = PV_cell
            Case "PMT"
                'I must be positive
                If i <= 0 Then
                    MsgBox "Cannot compute perpetuity if interest rate (I) is not positive.", vbCritical
                    Exit Sub
                End If
                If Mode = 0 Then
                    PMT_cell.Value = -PV * i
                Else
                    PMT_cell.Value = -PV * i / (1 + i)
                End If
                Set computedCell = PMT_cell
        End Select
    
    'N is a number (finite annuity)
    Else
        Select Case blankCellAddress
        
            Case "N"
                If i = 0 Then
                    If PMT = 0 Then
                        MsgBox "Cannot compute N if Interest Rate and Payment are both zero.", vbCritical
                        N_cell.Value = ""
                        Exit Sub
                    Else
                        N_cell.Value = -(PV + FV) / PMT
                    End If
                Else
                    N_cell.Value = Application.WorksheetFunction.NPer(i, PMT, PV, FV, Mode)
                End If
                Set computedCell = N_cell
                
            Case "I"
                I_cell.Value = Application.WorksheetFunction.Rate(N, PMT, PV, FV, Mode) * 100
                If I_cell.Value = "#NUM!" Then
                    MsgBox "Rate calculation failed. Check your inputs.", vbCritical
                    I_cell.Value = ""
                End If
                Set computedCell = I_cell
            
            Case "PV"
                PV_cell.Value = Application.WorksheetFunction.PV(i, N, PMT, FV, Mode)
                Set computedCell = PV_cell
                
            Case "PMT"
                PMT_cell.Value = Application.WorksheetFunction.PMT(i, N, PV, FV, Mode)
                Set computedCell = PMT_cell
                
            Case "FV"
                FV_cell.Value = Application.WorksheetFunction.FV(i, N, PMT, PV, Mode)
                Set computedCell = FV_cell
            
        End Select
    
    End If
    
    'Color the computed cell
    If Not computedCell Is Nothing Then
        computedCell.Interior.Color = vbGreen 'Apply light green color
        Set LastColoredCell = computedCell 'Store this cell for the next run
    End If
    
    Exit Sub
    
    
ErrorHandler:
        MsgBox "An error occured during calculation. Please check your inputs.", vbCritical
        'Clear the blank cell if an error occured to indicate failure
        Select Case blankCellAddress
            Case "N": N_cell.Value = ""
            Case "I": I_cell.Value = ""
            Case "PV": PV_cell.Value = ""
            Case "PMT": PMT_cell.Value = ""
            Case "FV": FV_cell.Value = ""
        Set LastColoredCell = Nothing
        End Select
        

End Sub
