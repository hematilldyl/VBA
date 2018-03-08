Attribute VB_Name = "Module4"
Function findMaxX()
    maxX = Worksheets("CRData").ChartObjects(1).Chart.Axes(xlCategory, xlPrimary).MaximumScale
    findMaxX = maxX
End Function

Sub ClearAll()

    With Worksheets("CRData")
        .Range("C2:K2").ClearContents
        .Range("C3:K3").ClearContents
        .Range("C4:K4").ClearContents
        .Range("C6").ClearContents
        .Range("I6:K6").ClearContents
        .Range("C14:C28").ClearContents
        .Range("D13:D29").ClearContents
        .Range("G14:G28").Value = False
        .Range("J14:J28").ClearContents
        .Range("K13:K29").ClearContents
        .Range("N14:N28").Value = False
        .Range("Q14:Q28").ClearContents
        .Range("R13:R29").ClearContents
        .Range("U14:U28").Value = False
        .Range("C5").ClearContents
        .Range("E6").ClearContents
    End With
    
End Sub

Sub GroupBox1_Click()
    Worksheets("CRData").Shapes("GroupBox1").Visible = False
End Sub

Sub GroupBox2_Click()
    Worksheets("CRData").Shapes("GroupBox2").Visible = False
End Sub

Sub GroupBox3_Click()
    Worksheets("CRData").Shapes("GroupBox3").Visible = False
End Sub

Sub RectNoCut1_Click()
    Worksheets("CRData").Shapes("RectNoCut1").Visible = False
End Sub

Sub RectNoCut2_Click()
    Worksheets("CRData").Shapes("RectNoCut2").Visible = False
End Sub

Sub RectNoCut3_Click()
    Worksheets("CRData").Shapes("RectNoCut3").Visible = False
End Sub

Sub RectSample2_Click()
    Worksheets("CRData").Range("N7").Value = True
    Worksheets("CRData").Shapes("RectSample2").Visible = False
End Sub

Sub RectSample3_Click()
    Worksheets("CRData").Range("U7").Value = True
    Worksheets("CRData").Shapes("RectSample3").Visible = False
End Sub

Sub cbInclude2_Change()
    If Worksheets("CRData").Range("N7") Then
        Worksheets("CRData").Shapes("RectSample2").Visible = False
    Else
        Worksheets("CRData").Shapes("RectSample2").Visible = True
    End If
End Sub

Sub cbInclude3_Change()
    If Worksheets("CRData").Range("U7") Then
        Worksheets("CRData").Shapes("RectSample3").Visible = False
    Else
        Worksheets("CRData").Shapes("RectSample3").Visible = True
    End If
End Sub

Sub updateChart2()

    If Worksheets("CRData").Range("N7") Then
        With Worksheets("CRData").ChartObjects(1).Chart
            With .SeriesCollection.NewSeries
                .Name = "Sample 2"
                .XValues = Worksheets("CRData").Range("J53:J67")
                .Values = Worksheets("CRData").Range("L53:L67")
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 8
                .MarkerForegroundColorIndex = 10 'green
                .MarkerBackgroundColorIndex = xlColorIndexNone
            End With
            With .SeriesCollection.NewSeries
                .Name = "Trend Line 2"
                .XValues = Worksheets("CRData").Range("J41:J51")
                .Values = Worksheets("CRData").Range("K41:K51")
                .Border.ColorIndex = 10 'green
                .Border.Weight = xlThick
                .Border.LineStyle = xlContinuous
                .MarkerStyle = xlNone
                .Smooth = True
            End With
            With .Legend
                For i = 1 To .LegendEntries.Count
                    If i = .LegendEntries.Count Then
                        .LegendEntries(i).Delete
                    End If
                Next
            End With
        End With
    Else
        With Worksheets("CRData").ChartObjects(1).Chart
            .SeriesCollection("Sample 2").Delete
            .SeriesCollection("Trend Line 2").Delete
        End With
    End If
    
End Sub

Sub updateChart3()
   
    If Worksheets("CRData").Range("U7") Then
        With Worksheets("CRData").ChartObjects(1).Chart
            With .SeriesCollection.NewSeries
                .Name = "Sample 3"
                .XValues = Worksheets("CRData").Range("Q53:Q67")
                .Values = Worksheets("CRData").Range("S53:S67")
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 8
                .MarkerForegroundColorIndex = 3 'red
                .MarkerBackgroundColorIndex = xlColorIndexNone
            End With
            With .SeriesCollection.NewSeries
                .Name = "Trend Line 3"
                .XValues = Worksheets("CRData").Range("Q41:Q51")
                .Values = Worksheets("CRData").Range("R41:R51")
                .Border.ColorIndex = 3 'red
                .Border.Weight = xlThick
                .Border.LineStyle = xlContinuous
                .MarkerStyle = xlNone
                .Smooth = True
            End With
            With .Legend
                For i = 1 To .LegendEntries.Count
                    If i = .LegendEntries.Count Then
                        .LegendEntries(i).Delete
                    End If
                Next
            End With
        End With
    Else
        With Worksheets("CRData").ChartObjects(1).Chart
            .SeriesCollection("Sample 3").Delete
            .SeriesCollection("Trend Line 3").Delete
        End With
    End If
End Sub



