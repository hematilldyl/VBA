Attribute VB_Name = "Module2"
Function InverseRegression(Y, X, y0, conf)
'===============================================================
' Code retrieved from 2016 ISEA Calculator template
'InverseRegression
'---------------------------------------------------------------
' Purpose: To perform least squares regression for independent X
'          and dependent Y and calculate the point estimate and
'          confidence interval for new value of the X given a
'          specific value of the dependent variable Y (inverse
'          regression).
'
' Author : Jeff Moreland
'          Ansell Healthcare Products, LLC.
'          511 Westinghouse Rd.
'          Pendleton, SC 29670
'          jmoreland@ansell.com
'
' Date   : 22 Aug 2009
'
' Notes  : Based on the Minitab macro CALIB.MAC (1).
' Refs   : [1] Derr, J. and S. Beder-Miller, CALIB.MAC. 1987,
'              Minitab, Inc. (http://www.minitab.com/en-US/
'              support/macros/default.aspx?action=code&id=11)
'          [2] Neter, J., et al., Applied linear statistical
'              models, ed., Irwin Homewood, IL: 1990.
'
'---------------------------------------------------------------
' Parameters
'   y, x, y0, conf
'-----------
'   y: Cell range for dependent variable (Y)
'   x: Cell range for independent variable (X)
'   y0: Y value for which X should be estimated
'   conf: Confidence level for interval estimate (0-100%)
'-----------
'
'---------------------------------------------------------------
' Returns
'-----------
'   An array containing the point estimate for x, the confidence
'   interval for x, the standard deviation of x, and a
'   correlation factor for the intervale (<0.1 is acceptable)
'---------------------------------------------------------------
'Revision History
'---------------------------------------------------------------
'
'===============================================================
    Dim intI As Integer
    
    'Perform intial regression using Linest, parameters:
    '   (1, 1): m       (1, 2): b
    '   (2, 1): se-m    (2, 2): se-b
    '   (3, 1): r2      (3, 2): se-y
    '   (4, 1): F       (4, 2): df
    '   (5, 1): ss-reg  (5, 2): ss-resid
    olsEstimates = Application.WorksheetFunction.LinEst(Y, X, True, True)
    m = olsEstimates(1, 1)
    b = olsEstimates(1, 2)
    df = olsEstimates(4, 2)
    n = df + 2
    ss_e = olsEstimates(5, 2)
    mse = ss_e / df
    r2 = olsEstimates(3, 1)
    x_bar = Application.WorksheetFunction.Average(X)
    
    'Calculate Sum of Sqaures for X-X_bar
    Dim x_dev() As Double
    ReDim x_dev(1 To n)
    
    intI = 1
    For Each v In X
        x_dev(intI) = v.Value - x_bar
        intI = intI + 1
    Next v
    
    Dim ss_x As Double
        ss_x = Application.WorksheetFunction.SumSq(x_dev)
        
    InverseRegression = calib(y0, m, b, n, df, mse, x_bar, ss_x, r2, 95)

End Function

Function calib(y0, m, b, n, df, mse, x_bar, ss_x, r2, conf)
'===============================================================
'calib
'---------------------------------------------------------------
' Purpose: To calculate the point estimate and confidence
'          interval for new value of the independent variable X
'          given a specific value of the dependent variable Y.
'
' Author : Jeff Moreland
'          Ansell Healthcare Products, LLC.
'          511 Westinghouse Rd.
'          Pendleton, SC 29670
'          jmoreland@ansell.com
'
' Date   : 22 Aug 2009
'
' Notes  : Based on the Minitab macro CALIB.MAC (1).
' Refs   : [1] Derr, J. and S. Beder-Miller, CALIB.MAC. 1987,
'              Minitab, Inc. (http://www.minitab.com/en-US/
'              support/macros/default.aspx?action=code&id=11)
'          [2] Neter, J., et al., Applied linear statistical
'              models, ed., Irwin Homewood, IL: 1990.
'
'---------------------------------------------------------------
' Parameters
'   y0, m, b, n, df, mse, x_bar, ss_x, conf
'-----------
'   y0: value of dependent variable Y for which the point
'       estimate and interval for X is desired
'   m: ols m of XvsY
'   b: ols b XvsY
'   n: number of observations
'   df: ols degrees of freedom (XvsY)
'   mse: ols mean sqaure error
'   x_bar: mean of X values
'   ss_x: sum of squares for X
'   conf: desired confidence level
'---------------------------------------------------------------
' Returns
'-----------
'   An array containing the slope, intercept, point estimate for
'   x, the confidence interval for x, the standard deviation of
'   x, and a correlation factor for the intervale (<0.1 is
'   acceptable)
'---------------------------------------------------------------
'Revision History
'---------------------------------------------------------------
'
'===============================================================
        Dim x_hat As Double     'Point estimate of X
        Dim clev As Double      'Confidence level
        Dim tabt As Double      't-value of area
        Dim s2 As Double        'variance of predicted x
        Dim s As Double         'standard deviation, sqrt(variance) for x
        Dim halfw As Double     'half width of the confidence interval
        Dim ci_low As Double    'lower bound of the condfidence interval
        Dim ci_high As Double   'upper bound of the confidence interval
        Dim width As Double     'width of the confidence interval
        Dim corr_fac As Double  'correlation factor
        
        x_hat = (y0 - b) / m
        clev = conf / 100
        tabt = Application.WorksheetFunction.TInv((1 - clev), df)
        s2 = (mse / (m ^ 2)) * (1 + (1 / n) + (((x_hat - x_bar) ^ 2) / ss_x))
        s = Sqr(s2)
        halfw = tabt * s
        ci_low = x_hat - halfw
        ci_high = x_hat + halfw
        width = 2 * halfw
        corr_fac = (tabt ^ 2) * mse / ((m ^ 2) * ss_x)
        calib = Array(m, b, x_hat, halfw, s, r2, corr_fac)
        
End Function



