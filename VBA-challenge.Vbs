Attribute VB_Name = "Module1"
Sub Tricker()
  'Loop Through all worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
    Dim WorksheetName As String
   
  'Count Last Row
    RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
  ' Set an initial variable for holding the Tricker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total Stock per Tricker
  Dim Total As Double
  Dim TotalStock As Double
  TotalStock = 0

  ' Keep track of the location for each Tricker in the summary table
  ws.Cells(1, 10).Value = "Ticker"
  ws.Cells(1, 11).Value = "Yearly Change"
  ws.Cells(1, 12).Value = "Percent Change"
  ws.Cells(1, 13).Value = "Total Stock Volume"
      
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
    
  
  ' Loop through all stocks
  For i = 2 To RowCount
  ClosePrice = Cells(i, 3).Value
 
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        OpenPrice = Cells(i, 3).Value
        
        End If
 
 ' Check if we are still within the same trikcer, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Tricker
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Yearly Change
      ClosePrice = Cells(i, 6).Value
      YearlyChange = ClosePrice - OpenPrice
      ' Add to the Percentage Change
      TotalStock = TotalStock + (Cells(i, 7).Value)
      PercentChange = FormatPercent(YearlyChange / OpenPrice)

      ' Print the Tricker Name in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Yearly & Percentage Change, Volume to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = YearlyChange
      ws.Range("L" & Summary_Table_Row).Value = PercentChange
      ws.Range("M" & Summary_Table_Row).Value = TotalStock

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Totals
      YearlyChange = 0
      TotalStock = 0

    ' If the cell immediately following a row is the same tricker...
    Else

      ' Add to the Total
      ClosePrice = Cells(i, 6).Value - ClosePrice
      TotalStock = TotalStock + Cells(i, 7).Value

    End If

    'Conditional Formatting positive change in green and negative change in red
     If Range("K" & Summary_Table_Row).Value <= 0 Then
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
     Else
      ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      End If
      
      
       '---------------------------------------------------
   'Bonus
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        Dim Max As Double
        Dim Min As Double
        Dim TSV As Double
        
    'Finding Greatest Increase,Decrease in Stock, also Greatest Volume
        Max = Application.WorksheetFunction.Max(Range("L:L").Value)
        Min = Application.WorksheetFunction.Min(Range("L:L").Value)
        TSV = Application.WorksheetFunction.Max(Range("M:M").Value)
        ws.Cells(2, 17).Value = FormatPercent(Max)
        ws.Cells(3, 17).Value = FormatPercent(Min)
        ws.Cells(4, 17).Value = TSV
      
     'Finding Tricker Greatest Increase,Decrease in Stock, also Greatest Volume
        ws.Range("P2").Value = Application.WorksheetFunction.Lookup(ws.Range("Q2").Value, ws.Range("L2:L3001").Value, ws.Range("J2:J3001").Value)
        ws.Range("P3").Value = Application.WorksheetFunction.Lookup(ws.Range("Q3").Value, ws.Range("L2:L3001").Value, ws.Range("J2:J3001").Value)
        ws.Range("P4").Value = Application.WorksheetFunction.Lookup(ws.Range("Q4").Value, ws.Range("L2:L3001").Value, ws.Range("J2:J3001").Value)
      
  Next i
  
    Next
     
End Sub



