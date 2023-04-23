Attribute VB_Name = "Module1"
Sub Dosomething()
'code sourced from https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
    Application.ScreenUpdating = True
End Sub

Sub stocks()

'declare variables
Dim ticker As String 'stock ticker
Dim netChange As Double 'yearly change from opening price at beginning of year to closing  price at end of year
Dim percentChange As Double 'percentage change from opening price at beginning of year to closing price at end of year
Dim totalVolume As Double 'total stock volume
Dim rowCounter As Double 'placeholder to loop through rows
Dim summaryTableRow As Integer 'placeholder to loop through rows
Dim opener As Double 'opening price for stock
Dim tickerMinMAX As String 'stock ticker for largest/smallest overall values
Dim maxPercentIncrease As Double 'greatest percent increase
Dim minPercentIncrease As Double 'greatest percent decrease
Dim maxTotalVolume As Double 'greatest total volume

summaryTableRow = 2 'skip the header (row 1) since we are populating that next

'populate headers of columns to output to
Range("I1").Value = "Ticker" 'populate header cell for stock ticker
Range("J1").Value = "Yearly Change" 'populate header for net change
Range("K1").Value = "Percent Change" 'populate header for percent change
Range("L1").Value = "Total Stock Volume" 'populate header for total volume

Range("O1").Value = "Ticker" 'populate header for stock ticker
Range("P1").Value = "Value" 'populate header for value

'populate row headers for table
Range("n2").Value = "Greatest % Increase" 'populate header for largest percent gain
Range("n3").Value = "Greatest % Decrease" 'populate header for largest percent decrease
Range("n4").Value = "Greatest Total Volume" 'populate header for largest total volume


netChange = 0 'initialize the yearly change, set to zero
rowCounter = 2 'initialize the proxy for iteration to skip header
opener = Cells(rowCounter, 3).Value 'initialize opening price

While IsEmpty(Cells(rowCounter, 1)) = False 'loop through all the rows until you reach an empty row
    If Cells(rowCounter + 1, 1).Value <> Cells(rowCounter, 1).Value Then 'if the next row contains a different stock
        'calculate final values for current stock
        ticker = Cells(rowCounter, 1).Value 'place value in ticker
        netChange = Cells(rowCounter, 6).Value - opener 'calculate the net change from opening price to closing price
        percentChange = netChange / opener 'calculate the percentage change of the opening price
        totalVolume = totalVolume + Cells(rowCounter, 7) 'sum up final total volume
    
        'populate summary table
        Range("I" & summaryTableRow).Value = ticker
        Range("J" & summaryTableRow).Value = netChange
        Range("K" & summaryTableRow).Value = FormatPercent(percentChange)
        Range("L" & summaryTableRow).Value = totalVolume
        
        'color code the summary table AND update max/min increases
        If netChange < 0 Then 'if change is negative color the cells red
            Range("J" & summaryTableRow).Interior.ColorIndex = 3
            Range("K" & summaryTableRow).Interior.ColorIndex = 3
            If percentChange < Range("P3").Value Then 'if this is a new minimum update table
                Range("P3").Value = FormatPercent(percentChange)
                Range("O3").Value = ticker
            End If
            
        Else 'if change is 0 or positive color the cells green
            Range("J" & summaryTableRow).Interior.ColorIndex = 4
            Range("K" & summaryTableRow).Interior.ColorIndex = 4
            If percentChange > Range("P2").Value Then 'if this is a new maximum update table
                Range("P2").Value = FormatPercent(percentChange)
                Range("O2").Value = ticker
            End If
                
        End If
        
        'update greatest total volume
        If totalVolume > Range("P4").Value Then 'if this is a new minimum update table
            Range("P4").Value = totalVolume
            Range("O4").Value = ticker
        End If
        
        summaryTableRow = summaryTableRow + 1 'move to next row in summary table
        
        'reset variable values for next iteration
        totalVolume = 0
        
        netChange = 0
        
        'proceed to the next row
        rowCounter = rowCounter + 1 'move "cursor" to next row
        opener = Cells(rowCounter, 3).Value 'set opening price for stock
        
    Else 'if next row contains same stock as current row
        totalVolume = totalVolume + Cells(rowCounter, 7) 'add up the total volume thus far
        rowCounter = rowCounter + 1 'proceed to next row
        
    End If
    
Wend

End Sub


