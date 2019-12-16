Attribute VB_Name = "Module1"
Sub stock_listing()

  ' Set an initial variable for the ticker
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per ticker
  Dim Ticker_Total As Double
  Ticker_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all ticker data
  For i = 2 To 1048576

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 3).Value

      ' Print the ticker in the Summary Table
      Range("G" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Ticker Amount to the Summary Table
      Range("H" & Summary_Table_Row).Value = Ticker_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Ticker Total
      Ticker_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Ticker Total
      Ticker_Total = Ticker_Total + Cells(i, 3).Value

    End If

  Next i

End Sub

