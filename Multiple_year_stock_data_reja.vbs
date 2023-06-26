Sub Stock()

'Define all  variables

Dim i, j, start As Integer

Dim OpeningPrice, ClosingPrice, YearlyChange, TotalStock, PercentChange As Double

Dim Ticker As String


'Looping through the worksheets

For Each ws In Worksheets

    'Assign column headers for Ticker, Yearly Change, Percent Change, & Total Stock Volume

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Assign integer for the loop to start
    start = 2
    NextTicker = 1
    TotalStock = 0

    'Determine the last row
    
    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'For each Ticker summrize and loop the yearly change, percent change, and total stock volume

        For i = 2 To EndRow

            'Get the Ticker from the column A

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value

            'To identify the next Ticker

            NextTicker = NextTicker + 1

            OpeningPrice = ws.Cells(NextTicker, 3).Value
            ClosingPrice = ws.Cells(i, 6).Value

            'Loop through to calculate Total Stock Volume from the volume column

            For j = NextTicker To i

                TotalStock = TotalStock + ws.Cells(j, 7).Value

            Next j

            'Loop through to calculate the Yearly Change and Percentage Change

            If OpeningPrice = 0 Then
                PercentChange = ClosingPrice

            Else
                YearlyChange = ClosingPrice - OpeningPrice
                PercentChange = YearlyChange / OpeningPrice

            End If

            'Output the results in the designated columns

            ws.Cells(start, 9).Value = Ticker
            ws.Cells(start, 10).Value = YearlyChange
            ws.Cells(start, 11).Value = PercentChange
            ws.Cells(start, 12).Value = TotalStock
            ws.Cells(start, 11).NumberFormat = "0.00%"

            'In the data summery when the first row task completed go to the next row

            start = start + 1

            'End the variable to zero & move i to NextTicker

            TotalStock = 0
            YearlyChange = 0
            PercentChange = 0
            NextTicker = i
            

        End If

    Next i

    ' Conditional formatting columns colors
            
        EndRow_YC = ws.Cells(Rows.Count, "J").End(xlUp).Row


            For j = 2 To EndRow_YC

            'if greater than or less than zero
                If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

                Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
            Next j
       
        
    'Get the Summary for Greatest % Increase, Greatest % Decrease, and Greatest Total Volume
  
    'Assign column and row header names
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
 
    
    'Get the maximum & minimum percentage from K column, & maximum total volume from L column
    ws.Range("R2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & EndRow)) * 100
    ws.Range("R3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & EndRow)) * 100
    ws.Range("R4") = WorksheetFunction.Max(ws.Range("L2:L" & EndRow))

    'Returns from the second row
    increase_greatest = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & EndRow)), ws.Range("K2:K" & EndRow), 0)
    decrease_greatest = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & EndRow)), ws.Range("K2:K" & EndRow), 0)
    volume_greatest = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & EndRow)), ws.Range("L2:L" & EndRow), 0)

    'Idetify ticker symbol for  total, greates increase, decrease, and volume
    ws.Range("Q2") = Cells(increase_greatest + 1, 9)
    ws.Range("Q3") = Cells(decrease_greatest + 1, 9)
    ws.Range("Q4") = Cells(volume_greatest + 1, 9)


Next ws

End Sub