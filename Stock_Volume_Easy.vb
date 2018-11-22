Sub Tarea2()

For Each ws In Worksheets

Dim ticker As String

  ' Set an initial variable for holding the total stock volume
  Dim Total_Stock As Double
  Total_Stock = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
'Table Headers:

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Total Stock Volume"

  'Find last row
  Dim lastrow As Long
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all stocks
  
  For i = 2 To lastrow

    ' Check if we are still within the same ticker...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker
      ticker = ws.Cells(i, 1).Value

      ' Add to the Total Stock volume
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = ticker

      ' Print the stock total to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Total_Stock

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total_Stock
      Total_Stock = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Total_Stock
      Total_Stock = Total_Stock + ws.Cells(i, 7).Value

    End If

  Next i

Next ws

End Sub
