Sub AllStocksAnalysisRefactored()

 Sheets("AllStocksAnalysis").Activate
    
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

    Range("A1").Value = "All Stocks (" + yearValue + ")"
'set up row headers

    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
        
'setup tickers array

    Dim tickers(11) As String
    
'Assign each ticker to element in array

    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    
'for loops

For i = 0 To 11
ticker = tickers(i)

    Sheets(yearValue).Activate
    rowStart = 2
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'reset total volume and number or rows for each ticker
    totalVolume = 0
    No_Of_Rows = 0

   For j = rowStart To rowEnd
   
        'increase totalVolume for each ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
                No_Of_Rows = No_Of_Rows + 1
            End If
            
        'set starting price
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                Dim startingPrice As Double
                startingPrice = Cells(j, 6).Value
            End If
            
        'set ending Price
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                Dim endPrice As Double
                endPrice = Cells(j, 6).Value
            End If

            
    Next j
      
'MsgBox No_Of_Rows

'write out total values in analysis worksheet
    Worksheets("AllStocksAnalysis").Activate
  
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = (endPrice / startingPrice) - 1


Next i

    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub


Sub formatAllStocksAnalysisTable()

    'formating
    Worksheets("AllStocksAnalysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i
    
End Sub

