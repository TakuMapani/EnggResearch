'Place Equation into a cell
Sub equationcell()
 
    Dim ws As Worksheet
    Dim wsNmae As String
    Dim wsStart As Worksheet
    Set wsStart = ActiveSheet
    Dim x As Integer
    Dim chrtObj As ChartObject
    Dim cht As Chart
    Dim ser As Series
    Dim serCol As SeriesCollection
    Dim AchartData As ChartObject
    'creating sheet to store equations
    Dim equationSheet As String
    Dim checkEquationSheet As String
    Dim strEquation As String
     Dim I As Integer
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        wsName = ws.Name
        Set chrtObj = Sheets(wsName).ChartObjects(1) ' Change to your sheet name here
        ws.Range("E1") = "Equation"
        With chrtObj.Chart
            For I = 1 To .SeriesCollection.Count
                If .SeriesCollection(I).Trendlines.Count > 0 Then
                    With .SeriesCollection(I).Trendlines(1)
                        If .DisplayEquation Then
                            Sheets(wsName).Range("E1").Offset(0, I).Value = .DataLabel.Text ' Change sheet name here as well
                        End If
                    End With
                End If
            Next I
        End With

    Next
End Sub




