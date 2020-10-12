Sub NewGraph()
    
    Dim ws As Worksheet
    Dim wsStart As Worksheet
    Set wsStart = ActiveSheet
    Dim x As Integer
    Dim chtObj As ChartObject
    Dim cht As Chart
    Dim ser As Series
    Dim serCol As SeriesCollection
    Dim AchartData As ChartObject
    'creating sheet to store equations
    Dim equationSheet As String
    Dim checkEquationSheet As String
    Dim strEquation As String
   
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        'filtering between -0.4V and 0.4V
        'Criteria1 and Criteria2 provide fields for limiting the values
        'With ws
        '    .Range("A1").AutoFilter field:=1, Criteria1:=">=-0.4", Operator:=xlAnd, Criteria2:="<=0.4"
        '
        'End With
        Rng = WorksheetFunction.CountA(Range("A1", Range("A1").End(xlDown)))
        
        'Delete all charts before creating new ones
        For Each AchartData In ws.ChartObjects
            AchartData.Delete
        Next
        
        'Change from milliAmps to nanoAmp
        'multiply Column B by 1e6 and store in column C
        Range("C1") = "nA"
        For x = 2 To Rng
            Range("C" & x).Formula = "=B" & x & "*1000000"
        Next
        
        
        'Function for ploting graph
        Set co = ws.ChartObjects.Add(Range("H5").Left, Range("H5").Top, 800, 600)
        Set cht = co.Chart
        
        With cht
            .HasTitle = True
            .ChartTitle.Text = ws.Name 'set chart name as sheetname
            Set serCol = .SeriesCollection
            Set ser = serCol.NewSeries
            
            With ser
                .Name = Range("C1").Value
                .XValues = Range(Range("A4"), Range("A4").End(xlDown))
                .Values = Range(Range("C4"), Range("C4").End(xlDown))
                .ChartType = xlXYScatterLinesNoMarkers
                .Trendlines.Add (xlLinear) 'adding trendline to the chart
            End With
            
            'displaying equation in trendline
            With ser.Trendlines(1)
                .DisplayEquation = True
            End With
            
            
        End With
        
        With cht.Axes(xlCategory)
            .HasTitle = True
            With .AxisTitle
                .Caption = "Voltage (V)"
            End With
            .TickLabelPosition = xlTickLabelPositionLow
        End With

        With cht.Axes(xlValue)
            .HasTitle = True
            With .AxisTitle
                .Caption = "Current (nA)"
            End With
            .TickLabelPosition = xlTickLabelPositionLow
        End With
 
    Next
        
        
End Sub
