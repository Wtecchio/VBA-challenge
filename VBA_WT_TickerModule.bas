Attribute VB_Name = "Module1"
Sub TickerTracker()


'My best soultion to running this subroutine on multiple sheets, running a for loop 3 times running the code 3 times

For i = 1 To 3

Sheets(i).Select



'marking the colums where the data analysis will go
Cells(1, 11).Value = "Ticker"
Cells(1, 12).Value = "Yearly Change"
Cells(1, 13).Value = "Percent Change"
Cells(1, 14).Value = "Total Stock Volume"


'Variable to hold Ticker symbol

tickerSymbol = ""

'Variable to hold the total stock vol of the ticker
totalVol = 0

' variable to hold the summary table starter row
SummaryTableRow = 2

'use function to find the last row in the sheet
 lastRow = Cells(Rows.Count, 1).End(xlUp).Row 'super important
 
 'loop from row 2 in column A out to the last row
 
 For Row = 2 To lastRow
 
    'check to see if the ticker changes
    
    
    If Cells(Row, 1).Value <> Cells(Row + 1, 1) Then 'This is the determine if there is a switch in tickers in the list by looking one row ahead compared to current row
    
        'if the ticker changes do...
        
        'first set the ticker brand
        
        tickerSymbol = Cells(Row, 1).Value
        
        'Save the last trade of the year
        
        lastClose = Cells(Row, 6).Value
        
        'Calculate the difference in the year from lastClose - firstOpen *ALSO CREATING VARIABLE "yearlyChange"
        
        yearlyChange = lastClose - fOpen
        
        
        'Calculate the percent change from the beg of the year to the end
        yearlyPercentChange = yearlyChange / fOpen
        
        
        
        'add the last charge from the row
        
        totalVol = totalVol + Cells(Row, 7).Value
        
        'add the ticker to the summary column in the summary table row
        Cells(SummaryTableRow, 11).Value = tickerSymbol
        
        'Add the difference from Open to close in the summary table
        Cells(SummaryTableRow, 12).Value = yearlyChange
        
        'Add the yearlyPercentChange to the summary table
        Cells(SummaryTableRow, 13).Value = yearlyPercentChange
        
        
        
        
        'Add the total vol to the column in the summary table row
        Cells(SummaryTableRow, 14).Value = totalVol
        

        
        
        'go to the next summary table row (add 1 on the value of the summary table)
        SummaryTableRow = SummaryTableRow + 1
        
        lastSummaryRow = SummaryTableRow - 1 'Storing this for next format loop
        
        ' reset the vol to 0
        totalVol = 0
        
        
    ElseIf Cells(Row, 1).Value <> Cells(Row - 1, 1) Then 'Will pull all opening prices of stocks
    
        fOpen = Cells(Row, 3).Value
    
    
    
    Else
            'if the ticker changes
            'add on to the total stockVol from the C column
        
        totalVol = totalVol + Cells(Row, 7).Value
            

        
    End If
    
    
    
        
 Next Row
 
 'Fromatting the numbers
    Columns(13).NumberFormat = "0.00%"
    
    
    
        

    
'Seperate Loop to format the data (possible to integrate on first loop, but would prefer to do multiple subroutines to be optimal
    
    
    
    
    For SummaryTableRow = 2 To lastSummaryRow

    
        If Cells(SummaryTableRow, 12).Value > 0 Then 'Green
        
            Cells(SummaryTableRow, 12).Interior.ColorIndex = 4
        
        ElseIf Cells(SummaryTableRow, 12).Value < 0 Then ' Red
            Cells(SummaryTableRow, 12).Interior.ColorIndex = 3
        
        Else 'gray
            Cells(SummaryTableRow, 12).Interior.ColorIndex = 15
        
        End If
        
        
    Next SummaryTableRow
    
    
'Find data for increase,decrease,deltaVol

gIncrease = 0
gIStock = ""

gDecrease = 0
gDStock = ""

gVol = 0
gVStock = ""


    For SummaryTableRow = 2 To lastSummaryRow
    
    
        'Increase
        If Cells(SummaryTableRow, 13) > gIncrease Then
        
        gIncrease = Cells(SummaryTableRow, 13)
        gIStock = Cells(SummaryTableRow, 11)
        
        Else
        
        End If
        
        'Decrease
        If Cells(SummaryTableRow, 13) < gDecrease Then
        
        gDecrease = Cells(SummaryTableRow, 13)
        gDStock = Cells(SummaryTableRow, 11)
        
        Else
        
        End If
        
        'Vol
        If Cells(SummaryTableRow, 14) > gVol Then
        
        gVol = Cells(SummaryTableRow, 14)
        gVStock = Cells(SummaryTableRow, 11)
        
        Else
        
        End If
    
    
   
    Next SummaryTableRow
    
    

 'Printing data
 
    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    
    
    'Increase
    Cells(2, 17).Value = gIStock
    Cells(2, 18).Value = gIncrease
    Cells(2, 18).NumberFormat = "0.00%"
    
    'Decrease
    Cells(3, 17).Value = gDStock
    Cells(3, 18).Value = gDecrease
    Cells(3, 18).NumberFormat = "0.00%"
    
    'Vol
    Cells(4, 17).Value = gVStock
    Cells(4, 18).Value = gVol
    Cells(4, 18).NumberFormat = "#,##0"





Columns.AutoFit

Next i


End Sub

