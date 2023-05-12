Attribute VB_Name = "Module1"
' Apply code to all worksheets and Auto-fit width of columns


Sub AllWorksheets()

    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Order
        
        Cells.EntireColumn.AutoFit
    
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
     Range("Q4").NumberFormat = General
    
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub Order()

'            Add new Columns and Rows names

            Range("J1").Value = "Ticker Symbol"
            Range("K1").Value = "Yearly Change "
            Range("L1").Value = "Percent Change "
            Range("M1").Value = "Total Stock Volume"
            
            Range("P1").Value = "Ticker Symbol"
            Range("Q1").Value = "Value "
            Range("O2").Value = "Greatest % Increase "
            Range("O3").Value = "Greatest % Decrease"
            Range("O4").Value = "Greatest Total Volume"

 'Find the last non-blank cell in a single row or column
            Dim lRow As Long
             
             'Find the last non-blank cell in column A(1)
            lRow = Cells(Rows.Count, 1).End(xlUp).Row
             
' Set Variable for holding the ticker symbol
Dim TickerSymbol As String

' Set Variable for holding the yearly change from the opening price at the beginning of a given year
' to the closing price at the end of that year.
Dim YearlyChange As Double
YearlyChange = 0
Columns("K:K").Select
    Selection.Style = "Currency"

' Set Variable for holding the percentage change from the opening price at the beginning of a given
' year to the closing price at the end of that year
Dim PercentChange As Double
PercentChange = 0
 Columns("L:L").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
' Set Variable for holding the total stock volume of the stock
Dim TotalStockVolume As Double
TotalStockVolume = 0

' Keep track of the location in the summary table
Dim SummaryTableRow As Integer
SummaryTableRow = 2

' Loop through all Ticker symbols
For i = 2 To lRow

    ' Check if we are still within the same ticker symbol, if it is not...
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Ticker Symbol
            TickerSymbol = Cells(i, 1).Value
    
            ' Print the Ticker Symbol in the Summary Table
            Range("J" & SummaryTableRow).Value = TickerSymbol
    
             ' Add to the Yearly Change
             YearlyChange = (Cells(i, 6).Value - Cells(i, 3).Value)
    
             ' Print the Yearly Change in the Summary Table
            Range("K" & SummaryTableRow).Value = YearlyChange
            
            ' Formatting
            If Cells(i, 11).Value < 0 Then
            Cells(i, 11).Interior.ColorIndex = 3
            
            Else
            Cells(i, 11).Interior.ColorIndex = 4
            
            End If
            
    ' Add to the Percent Change
    PercentChange = ((Cells(i, 6).Value) / (Cells(i, 3).Value)) / Cells(i, 3)
     
     ' Print the Percent Change in the Summary Table
      Range("L" & SummaryTableRow).Value = PercentChange
            
     ' Add to the Total Stock Volume
     TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
    ' Print the Total Stock Volume in the Summary Table
    Range("M" & SummaryTableRow).Value = TotalStockVolume

    ' Add one to the summary table row
    SummaryTableRow = SummaryTableRow + 1
    
            ' Reset the Yearly Change
            YearlyChange = 0
    
            '  Reset the Percent Change
            PercentChange = 0
         
            '  Reset the Total Stock Volume
            TotalStockVolume = 0
    
    
    ' If the cell inmediately following a row is the same Ticker Symbol...
    
    Else
    
    ' Add to the Yearly Change
    YearlyChange = YearlyChange + ((Cells(i, 6).Value - Cells(i, 3).Value) / Cells(i, 6))

    ' Add to the Percent Change
    PercentChange = (((Cells(i, 6).Value) / (Cells(i, 3).Value)) / 100)
    
    ' Add to the Total Stock Volume
     TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
    
    End If
    
Next i
    
' --------------------------------------------------



' Table 2 with Max and Min
'----------------------------


' Find last row in column J
Dim LastRowB As Long
LastRowB = Cells(Rows.Count, 10).End(xlUp).Row

' Variables
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVol As Double

For i = 2 To LastRowB

' Max Percent
        If Cells(i, 12).Value > MaxPercent Then
            MaxPercent = Cells(i, 12).Value
            Cells(2, 17) = MaxPercent
            Cells(2, 16).Value = Cells(i, 10).Value

        Else
        MaxPercent = MaxPercent

        End If
        
' Min Percent
        If Cells(i, 12).Value < MinPercent Then
            MinPercent = Cells(i, 12).Value
            Cells(3, 17) = MinPercent
            Cells(3, 16).Value = Cells(i, 10).Value

        Else
        MinPercent = MinPercent

        End If
         
 ' Max Volume
        If Cells(i, 13).Value > MaxVol Then
            MaxVol = Cells(i, 13).Value
            Cells(4, 17) = MaxVol
            Cells(4, 16).Value = Cells(i, 10).Value

        Else
        MaxVol = MaxVol

        End If
        
Next i
        
End Sub

