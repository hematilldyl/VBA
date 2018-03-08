Attribute VB_Name = "Module1"
Function CensoredRegression(Y, X, c)
'===============================================================
' Code retrieved from 2016 ISEA Calculator template
' CensoredRegression
'---------------------------------------------------------------
' Purpose: To detemine least squares regression estimates for
'          datasets that contain right-censored values based on
'          iterative procedure proposed by Schmee and Hahn(1).
'---------------------------------------------------------------
' Parameters
'-----------
'   y: Cell range of dependent variable (Y)
'   x: Cell range of independent vairbale (X)
'   c: Cell range of censors indicators
'      (0=Not Censored, 1=Censored)
'   conf: Confidence Level (0-100%)
'---------------------------------------------------------------
' Returns
'-----------
'   A Linest array with new paramter estimates
'===============================================================
    Dim x_r() As Double     'Array to hold x values
    Dim y_r() As Double     'Array to hold y values
    Dim y_0() As Double     'Array to hold original y values to keep track of censored values
    Dim c_r() As Integer    'Array to hold censor indicator values
    Dim oldSlope As Double 'Holds slope from previous iteration, used to check convergence
    Dim newSlope As Double 'Holds slope from current iteration, used to check convergence
    Dim convergeThreshold As Double  'Minimum difference between successive slopes to reach convergence
        convergeThreshold = 0.0001
    Dim maxIterations As Integer   'Maximum number of iterations
        maxIterations = 17
    Dim intI As Integer
    Dim intJ As Integer
    
    'Get values from worksheet ranges (for debugging)
'    Set x = Worksheets("CR").Range("D2:D41")
'    Set y = Worksheets("CR").Range("E2:E41")
'    Set c = Worksheets("CR").Range("F2:F41")
    
    'Set array holders to proper length
    intI = X.Count
    ReDim x_r(1 To intI)
    ReDim y_r(1 To intI)
    ReDim c_r(1 To intI)

    'Remove any empty elements in Ranges to accomodate Linest function
    'Based on empty X's only
    intI = 1
    intJ = 0
    For Each v In X
        If Not (IsEmpty(v)) Then
            x_r(intI) = v.Value
            y_r(intI) = Y.Cells(intI, 1).Value
            c_r(intI) = c.Cells(intI, 1).Value
            intJ = intJ + 1
        End If
        intI = intI + 1
    Next v

    ReDim Preserve x_r(1 To intJ)
    ReDim Preserve y_r(1 To intJ)
    ReDim Preserve c_r(1 To intJ)
    
    
    y_0 = y_r   'Preserve original Y values for later iterations

    'Perform intial regression using Linest, parameters:
    '   (1, 1): m       (1, 2): b
    '   (2, 1): se-m    (2, 2): se-b
    '   (3, 1): r2      (3, 2): se-y
    '   (4, 1): F       (4, 2): df
    '   (5, 1): ss-reg  (5, 2): ss-resid
    estimates = Application.WorksheetFunction.LinEst(y_r, x_r, True, True)
    oldSlope = estimates(1, 1)
    newSlope = oldSlope + 1
    
    'Enter estimates into spreadhseet (for debugging)
'    Worksheets("CR").Range("Q12").Select
'    ActiveCell.Value = oldSlope
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = estimates(1, 2)
'    ActiveCell.Offset(0, 1).Select
'    ActiveCell.Value = Sqr(estimates(5, 2) / estimates(4, 2))
'
    'Perform iterative method checking for convergence or until max
    'number of iterations is acheived
    intI = 0
    Do While hasNotConvergence(oldSlope, newSlope, convergeThreshold) Or intI < maxIterations
        oldSlope = newSlope
'        ActiveCell.Offset(1, 4).Select
        estimates = iterate(x_r, y_r, y_0, c_r, estimates)
        newSlope = estimates(1, 1)
'        ActiveCell.Offset(0, -10).Select
'        ActiveCell.Value = estimates(1, 1)
'        ActiveCell.Offset(0, 1).Select
'        ActiveCell.Value = estimates(1, 2)
'        ActiveCell.Offset(0, 1).Select
'        ActiveCell.Value = Sqr(estimates(5, 2) / estimates(4, 2))

        intI = intI + 1
    Loop
    
    CensoredRegression = estimates
    
End Function

Function iterate(x_r, y_r, y_0, c_r, estimates)
'===============================================================
'iterate
'---------------------------------------------------------------
' Purpose: Detrermine new values for the censored Y's based on
'          a log-normal distribution and estimate new lsr
'          parameters.
'---------------------------------------------------------------
' Parameters
'-----------
'   x_r: x value array
'   y_r: current y value array
'   y_0: original y value array
'   c_r: censor array (0=not censored, 1=censored)
'   estimates: linest estimates from previous iteration
'---------------------------------------------------------------
' Returns : linest parameter estimates for new y's.
'===============================================================

    Dim slope As Double
        slope = estimates(1, 1)
        'oldSlope = slope       'For convergence check
    Dim intI As Integer
    Dim df As Integer
        df = estimates(4, 2)
    Dim x_bar As Double
        x_bar = Application.WorksheetFunction.Average(x_r)
    Dim y_bar As Double
        y_bar = Application.WorksheetFunction.Average(y_r)
    Dim x_dev() As Double
        x_dev = x_r
        
    intI = 1
    For Each v In x_dev
        x_dev(intI) = v - x_bar
        intI = intI + 1
    Next v
    
    Dim ss_x As Double
        ss_x = Application.WorksheetFunction.SumSq(x_dev)
    Dim y_dev() As Double
        y_dev = y_r
        
    intI = 1
    For Each v In y_dev
        y_dev(intI) = v - y_bar
        intI = intI + 1
    Next v
    
    Dim ss_y As Double
        ss_y = Application.WorksheetFunction.SumSq(y_dev)
    Dim intercept As Double
        intercept = estimates(1, 2)
    Dim r_sq As Double
        r_sq = estimates(3, 1)
    Dim se As Double
        se = estimates(3, 2)
    Dim ss_total As Double
        ss_total = ss_y
    Dim y_hat() As Double
        y_hat = x_r
        
    intI = 1
    For Each v In y_hat
        y_hat(intI) = (slope * v) + intercept
        intI = intI + 1
    Next v
    
    Dim ss_e As Double
        ss_e = estimates(5, 2)
        'ss_e = Application.WorksheetFunction.SumXMY2(y, y_hat)
    Dim mse As Double
        mse = ss_e / df
    Dim std_dev As Double
        'std_dev = Sqr(e(5, 2) / df)
        'std_dev = Sqr(ss_e / df)
        'std_dev = Sqr(mse)
        std_dev = estimates(3, 2)
        
    Dim last_x As Double
        last_x = 0
    'Estimate new Y values is censored
    intI = 1
    For Each v In c_r
        If v > 0 Then   'v=1 means value is censored
            this_c = y_0(intI)  'Procedure uses intial censored Y values for determining log-normal parameters
            u0 = (slope * x_r(intI)) + intercept
            Z = (this_c - u0) / std_dev
            Pi = Application.WorksheetFunction.Pi()
            z_ord = (1 / Sqr(2 * Pi)) * Exp(-0.5 * (Z ^ 2)) 'Add ref
            z_area = Application.WorksheetFunction.NormSDist(Z)
            ua = u0 + (std_dev * (z_ord / (1 - z_area)))
            y_r(intI) = ua
            
'            If last_x <> x_r(intI) Then
'                last_x = x_r(intI)
'                ActiveCell.Offset(0, 1).Select
'                ActiveCell.Value = ua
'            End If
            
        End If
        
        intI = intI + 1
    Next v

    'Determine LSR estimates from new Y's
    iterate = Application.WorksheetFunction.LinEst(y_r, x_r, True, True)
        
End Function

Function hasNotConvergence(oldSlope, newSlope, threshold)
'===============================================================
'hasNotConvergence
'---------------------------------------------------------------
' Purpose: Detrermine if the difference between the current
'          slope and previous slope is less than a threshold.
'---------------------------------------------------------------
' Parameters
'-----------
'   threshold:  minimum slope difference to stop iteration
'---------------------------------------------------------------
' Returns : True when convergence has not been met.
'           False when converenge has been met.
'---------------------------------------------------------------
' Revision History
'---------------------------------------------------------------
'===============================================================
    slopeDif = Abs(oldSlope - newSlope)
    If slopeDif < threshold Then
        hasNotConvergence = False
    Else
        hasNotConvergence = True
    End If
    
End Function


