Sub Stockmarketvolume()

'declare variable ws as worksheet and loop through all sheets in the workbook

Dim ws As Worksheet

For Each ws In Worksheets

'print headers for the first results table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'print rows and column names for second results table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
'caculations for first results table

'declare variables to evaluate total volume and store ticker's name
    Dim totalvoulme As Long
    Dim ticker As String
    
'make the last row dynamic
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'define a variable that will help in writing the ticker and total volume into different rows in 1st results table
    Dim rowno As Long
    rowno = 2

'define the counter varaible

    Dim i As Long, j As Long
    
'define variables to calculate yearly change(closing value of stock - opening value of stock)
'define variable to find percent chage (yearly change/opening value of the stock)

    Dim openvalue As Double
    Dim closevalue As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    
'loop through all rows in this sheet using a conditional
'logic works when ticker names are sorted - approach the cellls of col1 from top to bottom
'use a conditional to check if ticker names are not the same


    For i = 2 To lastrow

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
            
            'If conditional is True, we perform a variety of tasks as under:
            
            'record ticker's name
            ticker = ws.Cells(i, 1).Value2
            
            'add the last volume of stocks for this ticker
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
            'Write the ticker name and total volume in the I and L columns
            'Use rowno to determine the row to which the value must be written
            ws.Range("I" & rowno).Value2 = ticker
            ws.Range("L" & rowno).Value2 = totalvolume
            
            'capture the close value corresponding to ith row (first day of the year)
            openvalue = ws.Cells(i - j, 3).Value2
            
            'capture open value corresponding to (i-j)th row (last day of the year)
            closevalue = ws.Cells(i, 6).Value2
            
            'find the yearly change and print it in the J column
            yearlychange = closevalue - openvalue
            ws.Range("J" & rowno).Value2 = yearlychange
            ws.Range("J" & rowno).NumberFormat = "0.00"
                
                'use conditional to highlight the values in the Yearly Change column
                'If values are less than zero then highlight in red(index = 3), else in green(index = 4).
                If ws.Range("J" & rowno).Value2 < 0 Then
                   ws.Range("J" & rowno).Interior.ColorIndex = "3"
                   Else: ws.Range("J" & rowno).Interior.ColorIndex = "4"
                End If
                
                'Debugging to check if there are values with denominator = 0
                If openvalue <> 0 Then
                    percentchange = (yearlychange / openvalue)
                    ws.Range("K" & rowno).Value = percentchange
                    ws.Range("K" & rowno).NumberFormat = "0.00%"
                    
                'Else: ws.Range("K" & rowno).Value2 = "divide by zero"
                End If
            
            'Preparation for calculations of next ticker - increment the row no by 1, both totalvolume and counter j are set to zero.
            rowno = rowno + 1
            totalvolume = 0
            j = 0
            
        Else
            'If conditional is False then add the values in Col7 to totalvolume
            totalvolume = totalvolume + ws.Cells(i, 7)
            'Use a counter to find the number of rows for each ticker. This counter helps in locating the index for opening value of stock
            j = j + 1

        End If

    Next i

'Calcluations for second results table

'find the lastrow of the first results table

Dim lastrow_results As Long
lastrow_results = ws.Cells(Rows.Count, 11).End(xlUp).Row

'declare varaibles to find the greatest % increse, greatest% decrease, greatest total volume and corresponding ticker names

Dim greatest_increase As Double, greatest_decrease As Double, greatest_totalvolume As Double
Dim greatest_increase_no As Double, greatest_decrease_no As Double, greatest_totalvolume_no As Double

'Debugging - MsgBox (lastrow_results)

'greatest % increase calculation using max worksheet functions.
'Match worksheet function is used to find the index of the row corresponding to the greatest % increase. This will be used to get the corresonding ticker's name
'write the values in the second results table
greatest_increase = WorksheetFunction.Max(ws.Range("K2:K" & lastrow_results))
greatest_increase_no = WorksheetFunction.Match(greatest_increase, ws.Range("K2:K" & lastrow_results), 0)
ws.Range("P2").Value = ws.Cells(greatest_increase_no + 1, 9)
ws.Range("Q2").Value = greatest_increase
ws.Range("Q2").NumberFormat = "0.00%"

'greatest % decrease calculation using min worksheet functions.
'Match worksheet function is used to find the index of the row corresponding to the greatest % decrease. This will give the corresponding ticker's name
'write the values in the second results table
greatest_decrease = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_results))
greatest_decrease_no = WorksheetFunction.Match(greatest_decrease, ws.Range("K2:K" & lastrow_results), 0)
ws.Range("P3").Value = ws.Cells(greatest_decrease_no + 1, 9)
ws.Range("Q3").Value = greatest_decrease
ws.Range("Q3").NumberFormat = "0.00%"

'greatest total volume loocation using the max worksheetfunction
'Match worksheet function is used to find the index of the row corresponding to the greatest total volume
'write the values in the second results table

greatest_totalvolume = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_results))
greatest_totalvolume_no = WorksheetFunction.Match(greatest_totalvolume, ws.Range("L2:L" & lastrow_results), 0)
ws.Range("P4").Value = ws.Cells(greatest_totalvolume_no + 1, 9)
ws.Range("Q4").Value = greatest_totalvolume

'autofit coulmns of result tables
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit

Next ws

End Sub