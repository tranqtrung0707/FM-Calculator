Attribute VB_Name = "ArithmeticAnnuity"
'Calculate time value of money variables for an arithmetic annuity


Function a_n(N As Variant, i As Double) As Double
' Returns a_angle_n, where i is in decimal form
    'Data validation
    If (N < 0) Or (i <= -100) Then
        'MsgBox "Invalid input for function a_n.", vbCritical
        Exit Function
    ElseIf N = "inf" Then 'perpetuity
        If (i <= 0) Then
            'MsgBox "Invalid input for function a_n. I must be positive for a perpetuity.", vbCritical
            Exit Function
        Else
            a_n = 1 / i
        End If
    Else 'finite annuity
        If (i = 0) Then
            a_n = N
        Else
            a_n = (1 - (1 + i) ^ (-N)) / i
        End If
    End If
End Function


Function Ia_n2(N As Variant, i As Double) As Double
' Returns the present value of a unit increasing annuity for (N-1) periods whose first payment is made at TIME 2, where i is in decimal form
    'Data validation
    If (N < 0) Or (i <= -100) Then
        MsgBox "Invalid input for function Ia_n2.", vbCritical
        Exit Function
    ElseIf N = "inf" Then 'perpetuity
        If (i <= 0) Then
            MsgBox "Invalid input for function Ia_n2. I must be positive for a perpetuity.", vbCritical
            Exit Function
        Else
            Ia_n2 = a_n(N, i) / i
        End If
    Else 'finite annuity
        If (i = 0) Then
            Ia_n2 = (N * (N - 1) / 2) / (1 + i) ^ 2
        Else
            Ia_n2 = (a_n(N, i) - N * (1 + i) ^ (-N)) / i
        End If
    End If
End Function



Sub ArithmeticAnnuity()

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
    Dim Mode_cell As range, N_cell As range, I_cell As range, PV_cell As range, P_cell As range, Q_cell As range, FV_cell As range

    'Assign the ranges to specific cells
    Set Mode_cell = ws.range("A_Mode_val")
    Set N_cell = ws.range("A_N_val")
    Set I_cell = ws.range("A_I_val")
    Set PV_cell = ws.range("A_PV_val")
    Set P_cell = ws.range("A_P_val")
    Set Q_cell = ws.range("A_Q_val")
    Set FV_cell = ws.range("A_FV_val")

    'Variables to hold the input and computed values
    Dim Mode As Integer, N As Variant 'N can be a number or "inf": denoting perpetuity
    Dim i As Double 'Interest rate will be in decimal form (e.g., 0.05 for 5%)
    Dim PV As Double, P As Double, Q As Double, FV As Double

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
    
    If IsEmpty(P_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "P"
    ElseIf Not IsNumeric(P_cell.Value) Then
        MsgBox "P must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        P = P_cell.Value
    End If
    
    If IsEmpty(Q_cell.Value) Then
        blankCount = blankCount + 1
        blankCellAddress = "Q"
    ElseIf Not IsNumeric(Q_cell.Value) Then
        MsgBox "Q must be a number. Please correct your input.", vbCritical
        Exit Sub
    Else
        Q = Q_cell.Value
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

    'MsgBox "N: " & N & vbCrLf & "I (decimal): " & I & vbCrLf & "PV: " & PV & vbCrLf & "PMT1: " & PMT1 & vbCrLf & "FV: " & FV & vbCrLf & vbCrLf & "Blank Count: " & blankCount & vbCrLf & "Blank Cell to Compute: " & blankCellAddress, vbInformation, "Debugging Input Values"

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
                'Solving i using the quadratic formula
                Dim Root1 As Double, Root2 As Double, RootMax As Double, RootMin As Double, Delta As Double
                'Root1 and Root2 are the two roots of the quadractic equation, RootMax is the larger one, Rootmin is the smaller one
                If Mode = 0 Then
                    Delta = P ^ 2 - 4 * PV * Q
                    If Delta < 0 Then
                        MsgBox "No solution for I", vbCritical
                        Exit Sub
                    Else
                        Root1 = (-P + Sqr(Delta)) / (2 * PV)
                        Root2 = (-P - Sqr(Delta)) / (2 * PV)
                        RootMax = Application.WorksheetFunction.Max(Root1, Root2)
                        RootMin = Application.WorksheetFunction.Min(Root1, Root2)
                        If RootMax <= 0 Then
                            MsgBox "No positive solution for I", vbCritical
                            Exit Sub
                        Else
                            I_cell.Value = RootMax * 100
                            If (RootMin > 0) And (RootMin < RootMax) Then
                                MsgBox "Another positive solution for I is: " & RootMin * 100 & "%"
                            End If
                        End If
                    End If
                Else
                    Delta = (P + Q) ^ 2 - 4 * (PV + P) * Q
                    If Delta < 0 Then
                        MsgBox "No solution for I", vbCritical
                        Exit Sub
                    Else
                        Root1 = (-(P + Q) + Sqr(Delta)) / (2 * (PV + P))
                        Root2 = (-(P + Q) - Sqr(Delta)) / (2 * (PV + P))
                        RootMax = Application.WorksheetFunction.Max(Root1, Root2)
                        RootMin = Application.WorksheetFunction.Min(Root1, Root2)
                        If RootMax <= 0 Then
                        MsgBox "No positive solution for I", vbCritical
                        Exit Sub
                        Else
                            I_cell.Value = RootMax * 100
                            If (RootMin > 0) And (RootMin < RootMax) Then
                                MsgBox "Another positive solution for I is: " & RootMin * 100 & "%"
                            End If
                        End If
                    End If
                End If
                Set computedCell = I_cell

            Case "PV"
                If Mode = 0 Then
                    PV_cell.Value = -P * a_n(N, i) - Q * Ia_n2(N, i)
                Else
                    PV_cell.Value = (-P * a_n(N, i) - Q * Ia_n2(N, i)) * (1 + i)
                End If
                Set computedCell = PV_cell
                
            Case "P"
                If Mode = 0 Then
                    P_cell.Value = -PV * i - Q / i
                Else
                    P_cell.Value = -PV / (1 + i) * i - Q / i
                End If
                Set computedCell = P_cell
                
            Case "Q"
                If Mode = 0 Then
                    Q_cell.Value = -PV * (i ^ 2) - P * i
                Else
                    Q_cell.Value = -PV / (1 + i) * (i ^ 2) - P * i
                End If
                Set computedCell = Q_cell


        End Select

    'N is a number (finite annuity)
    Else
        'Scratch worksheet
        Set scratchWs = ThisWorkbook.Sheets("Scratch Works (DO NOT DELETE)")
        
        Select Case blankCellAddress

            Case "N"
                'In order to use goal seek, we must first copy our data to another sheet
                ws.range("A_").Copy Destination:=scratchWs.range("K2")
                'Then calculate the PV cell for a given N
                If Mode = 0 Then
                    scratchWs.range("K5").Formula = "=-K6*a_n(K3,K4/100) - K7*Ia_n2(K3, K4/100) - K8*(1+K4/100)^(-K3)"
                Else
                    scratchWs.range("K5").Formula = "=-K6*a_n(K3,K4/100)*(1+K4/100) - K7*Ia_n2(K3, K4/100)*(1+K4/100) - K8*(1+K4/100)^(-K3)"
                End If
                'Use Goal Seek to find the value of N that makes the value in the PV cell equals PV, initial value N=1
                scratchWs.range("K3").Value = 1
                scratchWs.range("K5").GoalSeek Goal:=PV, ChangingCell:=scratchWs.range("K3")
                'Copy N from Sheet2 to Sheet1
                scratchWs.range("K3").Copy Destination:=N_cell
                Set computedCell = N_cell

            Case "I"
                'In order to use goal seek, we must first copy our data to another sheet
                ws.range("A_").Copy Destination:=scratchWs.range("K2")
                'Then calculate the PV cell for a given N
                If Mode = 0 Then
                    scratchWs.range("K5").Formula = "=-K6*a_n(K3,K4/100) - K7*Ia_n2(K3, K4/100) - K8*(1+K4/100)^(-K3)"
                Else
                    scratchWs.range("K5").Formula = "=-K6*a_n(K3,K4/100)*(1+K4/100) - K7*Ia_n2(K3, K4/100)*(1+K4/100) - K8*(1+K4/100)^(-K3)"
                End If
                'Use Goal Seek to find the value of N that makes the value in the PV cell equals PV, initial value I=10(%)
                scratchWs.range("K4").Value = 10
                scratchWs.range("K5").GoalSeek Goal:=PV, ChangingCell:=scratchWs.range("K4")
                'Copy N from Sheet2 to Sheet1
                scratchWs.range("K4").Copy Destination:=I_cell
                Set computedCell = I_cell

            Case "PV"
                If Mode = 0 Then
                    PV_cell.Value = -P * a_n(N, i) - Q * Ia_n2(N, i) - FV * (1 + i) ^ (-N)
                Else
                    PV_cell.Value = -P * a_n(N, i) * (1 + i) - Q * Ia_n2(N, i) * (1 + i) - FV * (1 + i) ^ (-N)
                End If
                Set computedCell = PV_cell

            Case "P"
                If Mode = 0 Then
                    P_cell.Value = -(PV + Q * Ia_n2(N, i) + FV * (1 + i) ^ (-N)) / a_n(N, i)
                Else
                    P_cell.Value = -(PV + Q * Ia_n2(N, i) * (1 + i) + FV * (1 + i) ^ (-N)) / (a_n(N, i) * (1 + i))
                End If
                Set computedCell = P_cell
                
            Case "Q"
                If Mode = 0 Then
                    Q_cell.Value = -(PV + P * a_n(N, i) + FV * (1 + i) ^ (-N)) / Ia_n2(N, i)
                Else
                    Q_cell.Value = -(PV + P * a_n(N, i) * (1 + i) + FV * (1 + i) ^ (-N)) / (Ia_n2(N, i) * (1 + i))
                End If
                Set computedCell = Q_cell
            
            Case "FV"
                If Mode = 0 Then
                    FV_cell.Value = -(PV + P * a_n(N, i) + Q * Ia_n2(N, i)) * (1 + i) ^ N
                Else
                    FV_cell.Value = -(PV + P * a_n(N, i) * (1 + i) + Q * Ia_n2(N, i) * (1 + i)) * (1 + i) ^ N
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
            Case "PV": PV_cell.Value = ""
            Case "P": P_cell.Value = ""
            Case "Q": Q_cell.Value = ""
            Case "FV": FV_cell.Value = ""
        Set LastColoredCell = Nothing
        End Select


End Sub


