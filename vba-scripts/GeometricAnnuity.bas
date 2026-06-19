Attribute VB_Name = "GeometricAnnuity"
'Calculate time value of money variables for a geometric annuity

' Returns a_angle_n_superscript_g, where i and g are in decimal form
Function a_n_g(N As Variant, i As Double, G As Double) As Double
    'Data validation
    If (N < 0) Or (i <= -100) Or (G <= -100) Then
        'MsgBox "Invalid input for function a_n_g."
        Exit Function
    ElseIf (i = G) Then
        a_n_g = 1 / (1 + i) * N
    Else
        a_n_g = ((1 - ((1 + G) / (1 + i)) ^ N) / (i - G))
    End If
End Function

'Returns a_angle_inf_superscript_g, where i and g are in decimal form
Function a_inf_g(i As Double, G As Double) As Double
    'Data Validation
    If (N < 0) Or (i <= -100) Or (G <= -100) Then
        'MsgBox "Invalid input for function a_inf_g."
        Exit Function
    ElseIf (i <= G) Then
        MsgBox "The perpetuity cannot be calculated for I <= G. Please check your input."
        Exit Function
    Else
        a_inf_g = 1 / (i - G)
    End If
End Function



Sub GeometricAnnuity()

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
    Dim Mode_cell As range, N_cell As range, I_cell As range, G_cell As range, PV_cell As range, PMT1_cell As range, FV_cell As range
    
    'Assign the ranges to specific cells
    Set Mode_cell = ws.range("G_Mode_val")
    Set N_cell = ws.range("G_N_val")
    Set I_cell = ws.range("G_I_val")
    Set G_cell = ws.range("G_G_val")
    Set PV_cell = ws.range("G_PV_val")
    Set PMT1_cell = ws.range("G_PMT1_val")
    Set FV_cell = ws.range("G_FV_val")
    
    'Variables to hold the input and computed values
    Dim Mode As Integer, N As Variant 'N can be a number or "inf": denoting perpetuity
    Dim i As Double, G As Double 'Interest rate and growth rate will be in decimal form (e.g., 0.05 for 5%)
    Dim PV As Double, PMT1 As Double, FV As Double
    
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
    
    If IsEmpty(G_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "G"
    ElseIf (Not IsNumeric(G_cell.Value)) Or (G_cell.Value < -100) Then
        MsgBox "G must be a number greater than or equal to -100. Please correct your input.", vbCritical
        Exit Sub
    Else
        G = G_cell.Value / 100 'Convert percentage to decimal for calculation
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

    If IsEmpty(PMT1_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "PMT1"
    ElseIf Not IsNumeric(PMT1_cell.Value) Then
        MsgBox "PMT1 must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        PMT1 = PMT1_cell.Value
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
            MsgBox "Cannot Compute FV for a perpetuity. Please leave I, G, PV, or PMT1 blank.", vbCritical
            Exit Sub
        End If

        'Perform perpetuity calculations
        Select Case blankCellAddress
        
            Case "I"
                'PV and PMT1 must have the opposite sign and PV<>0
                If (PMT1 * PV > 0) Or PV = 0 Then
                    MsgBox "Cannot compute I with the given PMT1 and PV.", vbCritical
                    Exit Sub
                End If
                If Mode = 0 Then
                    I_cell.Value = ((PMT1 / -PV) + G) * 100
                Else
                    If (PV + PMT1 <= 0) Then
                        MsgBox "Cannot compute I with the given PMT1 and PV.", vbCritical
                        Exit Sub
                    Else
                        I_cell.Value = (PV * G - PMT1) / (PV + PMT1) * 100
                    End If
                End If
                Set computedCell = I_cell
                
            Case "G"
                'PV and PMT1 must have the opposite sign and PV<>0
                If (PMT1 * PV > 0) Or PV = 0 Then
                    MsgBox "Cannot compute I with the given PMT1 and PV.", vbCritical
                    Exit Sub
                End If
                If Mode = 0 Then
                    If (PV + PMT1 <= 0) Then
                        MsgBox "Cannot compute I with the given PMT1 and PV.", vbCritical
                        Exit Sub
                    Else
                        G_cell.Value = (i - (PMT1 / -PV)) * 100
                    End If
                Else
                    G_cell.Value = (i - (PMT1 * (1 + i) / -PV)) * 100
                End If
                Set computedCell = G_cell
                
            Case "PV"
                If Mode = 0 Then
                    PV_cell.Value = -PMT1 * a_inf_g(i, G)
                Else
                    PV_cell.Value = -PMT1 * a_inf_g(i, G) * (1 + i)
                End If
                Set computedCell = PV_cell
                
            Case "PMT1"
                If Mode = 0 Then
                    PMT1_cell.Value = -PV / a_inf_g(i, G)
                Else
                    PMT1_cell.Value = -PV / (1 + i) / a_inf_g(i, G)
                End If
                Set computedCell = PMT1_cell
                
        End Select
    
    'N is a number (finite annuity)
    Else
        'Scratch worksheet
        Set scratchWs = ThisWorkbook.Sheets("Scratch Works (DO NOT DELETE)")
                
        Select Case blankCellAddress

            Case "N"
                'In order to use goal seek, we must first copy our data to another sheet
                ws.range("G_").Copy Destination:=scratchWs.range("G2")
                'Then calculate the PV cell for a given N
                If Mode = 0 Then
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100) - G8*(1+G4/100)^(-G3)"
                Else
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100)*(1+G4/100) - G8*(1+G4/100)^(-G3)"
                End If
                'Use Goal Seek to find the value of N that makes the value in the PV cell equals PV, initial value = 1
                scratchWs.range("G3").Value = 1
                scratchWs.range("G6").GoalSeek Goal:=PV, ChangingCell:=scratchWs.range("G3")
                'Copy N from Sheet2 to Sheet1
                scratchWs.range("G3").Copy Destination:=N_cell
                Set computedCell = N_cell

            Case "I"
                'In order to use goal seek, we must first copy our data to another sheet
                ws.range("G_").Copy Destination:=scratchWs.range("G2")
                'Then calculate the PV cell for a given I
                If Mode = 0 Then
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100) - G8*(1+G4/100)^(-G3)"
                Else
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100)*(1+G4/100) - G8*(1+G4/100)^(-G3)"
                End If
                'Use Goal Seek to find the value of I that makes the value in the PV cell equals PV, initial value for I = G+1
                scratchWs.range("G4").Value = 100 * G + 1
                scratchWs.range("G6").GoalSeek Goal:=PV, ChangingCell:=scratchWs.range("G4")
                'Copy I from Sheet2 to Sheet1
                scratchWs.range("G4").Copy Destination:=I_cell
                Set computedCell = I_cell
                
            Case "G"
                'In order to use goal seek, we must first copy our data to another sheet
                ws.range("G_").Copy Destination:=scratchWs.range("G2")
                'Then calculate the PV cell for a given I
                If Mode = 0 Then
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100) - G8*(1+G4/100)^(-G3)"
                Else
                    scratchWs.range("G6").Formula = "=-G7*a_n_g(G3,G4/100,G5/100)*(1+G4/100) - G8*(1+G4/100)^(-G3)"
                End If
                'Use Goal Seek to find the value of I that makes the value in the PV cell equals PV, initial value for G=I-1
                scratchWs.range("G5").Value = 100 * i - 1
                scratchWs.range("G6").GoalSeek Goal:=PV, ChangingCell:=scratchWs.range("G5")
                'Copy I from Sheet2 to Sheet1
                scratchWs.range("G5").Copy Destination:=G_cell
                Set computedCell = G_cell

            Case "PV"
                If Mode = 0 Then
                    PV_cell.Value = -PMT1 * a_n_g(N, i, G) - FV * (1 + i) ^ (-N)
                Else
                    PV_cell.Value = -PMT1 * a_n_g(N, i, G) * (1 + i) - FV * (1 + i) ^ (-N)
                End If
                Set computedCell = PV_cell

            Case "PMT1"
                If Mode = 0 Then
                    PMT1_cell.Value = -(PV + FV * (1 + i) ^ (-N)) / a_n_g(N, i, G)
                Else
                    PMT1_cell.Value = -(PV + FV * (1 + i) ^ (-N)) / (a_n_g(N, i, G)) / (1 + i)
                End If
                Set computedCell = PMT1_cell

            Case "FV"
                If Mode = 0 Then
                    FV_cell.Value = -(PV + PMT1 * a_n_g(N, i, G)) * (1 + i) ^ N
                Else
                    FV_cell.Value = -(PV + PMT1 * a_n_g(N, i, G) * (1 + i)) * (1 + i) ^ N
                End If
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
            Case "G": G_cell.Value = ""
            Case "PV": PV_cell.Value = ""
            Case "PMT1": PMT1_cell.Value = ""
            Case "FV": FV_cell.Value = ""
        Set LastColoredCell = Nothing
        End Select
        

End Sub

