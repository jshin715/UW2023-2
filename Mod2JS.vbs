'AAB : 455750984
'AAF : 2080570716
'BAJQ: 85480670
'BAOV  : 832425430

Sub stock_data_h()

    'Define variables
    Dim Ticker As String
    Dim Opening_Price, Closing_Price, Yearly_Change, Percent_Change, Daily_Volume, Total_Stock_Volume As Double
    Dim Greatest_Percent_Increase, Greatest_Percent_Decrease, Greatest_Total_Volume As Double
    Dim Summary_Table_Row As Integer
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        ws.Activate
        'Input headers for output
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"



    
        Total_Stock_Volume = 0

        ' Keep track of the location for each stock in the summary table
        Summary_Table_Row = 2

        'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            ' Check if we are still within the same stock
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' Set the Ticker name
                Ticker = Cells(i, 1).Value

                ' Add to the Volume Total
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
                ' Print the Credit Card Brand in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker
    
                ' Print the Brand Amount to the Summary Table
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
          
                ' Reset the Brand Total
                Total_Stock_Volume = 0
         
                ' If the cell immediately following a row is the same...
            Else

                ' Add to the Brand Total
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            End If
    
         Next i
         
         Summary_Table_Row = 2
        Opening_Price = 0
         For i = 2 To lastrow

            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                Closing_Price = Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                Percent_Change = (Closing_Price - Opening_Price) / Opening_Price
                Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)
                Summary_Table_Row = Summary_Table_Row + 1
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            
                Opening_Price = Cells(i, 3).Value
                
            End If

    
        Next i
        
        Greatest_Percent_Increase = WorksheetFunction.Max(Columns("k:k"))
        Greatest_Percent_Decrease = WorksheetFunction.Min(Columns("k:k"))
        Greatest_Total_Volume = WorksheetFunction.Max(Columns("l"))
        
        For i = 2 To lastrow
            If Greatest_Percent_Increase = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf Greatest_Percent_Decrease = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            End If
        Next i
        
        For i = 2 To lastrow
            If Greatest_Total_Volume = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        Range("Q2").Value = FormatPercent(Greatest_Percent_Increase)
        Range("Q3").Value = FormatPercent(Greatest_Percent_Decrease)
        Range("Q4").Value = Greatest_Total_Volume
        
        For i = 2 To lastrow
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        ws.Columns("A:Q").AutoFit
    
    Next ws

End Sub



