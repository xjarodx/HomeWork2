Sub Summmary()

For Each ws In Worksheets
  Dim WorksheetName As String
  WorksheetName = ws.Name
  
  ' Set an initial variable for holding the ticker
  Dim Ticker As String
  
  ' Set an initial variable for holding the total per credit card brand
  Dim Ticker_Volume As Variant
  Ticker_Volume = 0

  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  'Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
  
  ws.Cells(1, 9) = "Ticker"
  ws.Cells(1, 10) = "Total Volume"
  
  ' Loop through all data
  For i = 2 To lastrow

    ' Check if same Ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker
      Ticker = ws.Cells(i, 1).Value

      ' Add to the Ticker Volume
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Ticker Volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Volume = 0
        
    Else
      ' Add to the Ticker Volume
      Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value

    End If

  Next i
  
Next ws

End Sub
