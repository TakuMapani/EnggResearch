Sub CombineTextFiles()
'updateby Extendoffice 20151015
    Dim xFilesToOpen As Variant
    Dim I As Integer
    Dim xWb As Workbook
    Dim xTempWb As Workbook
    Dim xDelimiter As String
    Dim xScreen As Boolean
    On Error GoTo ErrHandler
    xScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    xDelimiter = "|"
    xFilesToOpen = Application.GetOpenFilename("Text Files (*.txt), *.txt", , "Kutools for Excel", , True)
    If TypeName(xFilesToOpen) = "Boolean" Then
        MsgBox "No files were selected", , "Kutools for Excel"
        GoTo ExitHandler
    End If
    I = 1
    Set xTempWb = Workbooks.Open(xFilesToOpen(I))
    xTempWb.Sheets(1).Copy
    Set xWb = Application.ActiveWorkbook
    xTempWb.Close False
    xWb.Worksheets(I).Columns("A:A").TextToColumns _
      Destination:=Range("A1"), DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=False, Semicolon:=False, _
      Comma:=False, Space:=False, _
      Other:=True, OtherChar:="|"
    Do While I < UBound(xFilesToOpen)
        I = I + 1
        Set xTempWb = Workbooks.Open(xFilesToOpen(I))
        With xWb
            xTempWb.Sheets(1).Move after:=.Sheets(.Sheets.Count)
            .Worksheets(I).Columns("A:A").TextToColumns _
              Destination:=Range("A1"), DataType:=xlDelimited, _
              TextQualifier:=xlDoubleQuote, _
              ConsecutiveDelimiter:=False, _
              Tab:=False, Semicolon:=False, _
              Comma:=False, Space:=False, _
              Other:=True, OtherChar:=xDelimiter
        End With
    Loop
ExitHandler:
    Application.ScreenUpdating = xScreen
    Set xWb = Nothing
    Set xTempWb = Nothing
    Exit Sub
ErrHandler:
    MsgBox Err.Description, , "Kutools for Excel"
    Resume ExitHandler
End Sub