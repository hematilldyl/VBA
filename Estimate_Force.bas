Attribute VB_Name = "Module3"
Sub testRoutine()
    Dim Y As Range
    Dim X As Range
    Dim c As Range
    
    Set Y = Worksheets("Sheet1").Range("F11:F25")
    'Set y = Application.WorksheetFunction.Log10(Worksheets("Sheet1").Range("C3:C17"))
    'logy = Application.WorksheetFunction.Log10(y)
    Set X = Worksheets("Sheet1").Range("C11:C25")
    Set c = Worksheets("Sheet1").Range("H11:H25")
    
    theseEstimates = EstimateRatingForce(Y, X, c, 20, 95)
    Debug.Print theseEstimates(0)
    
End Sub
Function EstimateRatingForce(Y, X, c, y0, conf)

    olsEstimates = CensoredRegression(Y, X, c)
    '   (1, 1): m       (1, 2): b
    '   (2, 1): se-m    (2, 2): se-b
    '   (3, 1): r2      (3, 2): se-y
    '   (4, 1): F       (4, 2): df
    '   (5, 1): ss-reg  (5, 2): ss-resid
    m = olsEstimates(1, 1)
    b = olsEstimates(1, 2)
    df = olsEstimates(4, 2)
    n = df + 2
    ss_e = olsEstimates(5, 2)
    r2 = olsEstimates(3, 1)
    mse = ss_e / df
    x_bar = Application.WorksheetFunction.Average(X)
    
    'Calculate Sum of Sqaures for X-X_bar
    Dim x_dev() As Double
    ReDim x_dev(1 To n)
    
    intI = 1
    Do While intI < n + 1
        x_dev(intI) = X(intI) - x_bar
        intI = intI + 1
    Loop
    
    Dim ss_x As Double
        ss_x = Application.WorksheetFunction.SumSq(x_dev)
        
    EstimateRatingForce = calib(y0, m, b, n, df, mse, x_bar, ss_x, r2, conf)

End Function


