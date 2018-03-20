'===============================================================
'---------------------------------------------------------------

' Purpose: Iteratively perform Simplex Algorithm on a transposed 
'          system of equations to minimize under contraints
'          generated from exponential smoothing forecasts and 
'          maintenance reqs for t number of forecast periods

'---------------------------------------------------------------
'---------------------------------------------------------------

' Author: Dylan Hematillake
' Date: 2018-03-20
' Version: 3

'---------------------------------------------------------------
'===============================================================


Attribute VB_Name = "Module1"
Sub Simplex()

Dim i, j, k, t As Integer
Dim hold As Double
Dim Dict(8) As String
Dim ws As Worksheet
Dim wb As Workbook

Dict(1) = "A"
Dict(2) = "B"
Dict(3) = "C"
Dict(4) = "D"
Dict(5) = "E"
Dict(6) = "F"
Dict(7) = "G"
Dict(8) = "H"

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Sheet1")
Set ws2 = wb.Sheets("Data")
ws.Activate


For t = 0 To 7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
hold = Range("C" & (t * 6 + 5)).Value
For Each c In Range("A" & (t * 6 + 5) & ":" & "H" & (t * 6 + 5))
    c.Value = c.Value / hold
Next c

i = 1
hold = Range("C" & (t * 6 + 6)).Value
For Each d In Range("A" & (t * 6 + 6) & ":" & "H" & (t * 6 + 6))
    d.Value = d.Value - hold * Range(Dict(i) & (t * 6 + 5)).Value
    i = i + 1
Next d
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
hold = Range("B" & (t * 6 + 4)).Value
For Each c In Range("A" & (t * 6 + 4) & ":" & "H" & (t * 6 + 4))
    c.Value = c.Value / hold
Next c

i = 1
hold = Range("B" & (t * 6 + 6)).Value
For Each d In Range("A" & (t * 6 + 6) & ":" & "H" & (t * 6 + 6))
    d.Value = d.Value - hold * Range(Dict(i) & 4).Value
    i = i + 1
Next d
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
hold = Range("A" & (t * 6 + 3)).Value
For Each c In Range("A" & (t * 6 + 3) & ":" & "H" & (t * 6 + 3))
    c.Value = c.Value / hold
Next c

i = 1
hold = Range("A" & (t * 6 + 6)).Value
For Each d In Range("A" & (t * 6 + 6) & ":" & "H" & (t * 6 + 6))
    d.Value = d.Value - hold * Range(Dict(i) & 3).Value
    i = i + 1
Next d

Next t

ws2.Activate
End Sub
