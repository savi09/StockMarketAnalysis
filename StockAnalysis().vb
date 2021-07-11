Sub StockAnalysis()

    Dim ws As Worksheet

'Loop through all sheets
    For Each ws In Worksheets
    
        'Define variable
            Dim ticker As String
            Dim totalOpen As Double
            Dim totalClose As Double
            Dim yearlyChange As Double
            Dim percentCh As Double
            Dim totalVol As Double
            Dim position As Integer
        
        'Top headers
            ws.Range("J1") = "Ticker"
            'Used the following two for testing purposes
            'ws.Range("K1") = "Total Open"
            'ws.Range("L1") = "Total Close"
            ws.Range("K1") = "Yearly Change"
            ws.Range("L1") = "Percentage Change"
            'Rounds column N
            ws.Columns("L").NumberFormat = Round(percentCh, 2)
            'Format Column N as a percentage
            ws.Columns("L").NumberFormat = "##.##%"
            ws.Range("M1") = "Total Stock Volume"
    
        'Reset Variables
            totalOpen = 0
            totalClose = 0
            yearlyChange = 0
            percentCh = 0
            totalVol = 0
            position = 2
        'Determine the last row
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Loop through rows
            For i = 2 To LastRow
        
            'If statement to group by ticker type
                If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                    'Adds last value of ticker
                    ticker = ws.Cells(i, 1).Value
                    totalOpen = totalOpen + ws.Cells(i, 3).Value
                    totalClose = totalClose + ws.Cells(i, 6).Value
                    yearlyChange = totalClose - totalOpen
                    'Avoids dividing by zero
                    If totalOpen = 0 Then
                        percentCh = 0
                        ws.Range("L" & position).NumberFormat = "General"
                    Else
                        percentCh = (yearlyChange / totalOpen)
                    End If
                    totalVol = totalVol + ws.Cells(i, 7).Value
                    ws.Range("J" & position).Value = ticker
                    'ws.Range("K" & position).Value = totalOpen
                    'ws.Range("L" & position).Value = totalClose
                    ws.Range("K" & position).Value = yearlyChange
                    ws.Range("L" & position).Value = percentCh
                    ws.Range("M" & position).Value = totalVol
                        
                        'Formatting
                        If (ws.Range("L" & position).Value >= 0) Then
                            ws.Range("L" & position).Interior.Color = vbGreen
                        Else
                            ws.Range("L" & position).Interior.Color = vbRed
                        End If
                        
                    'Move to the next row
                    position = position + 1
                    
                    'Reset variables
                    ticker = 0
                    totalOpen = 0
                    totalClose = 0
                    yearlyChange = 0
                    percentCh = 0
                    totalVol = 0
            
                Else
                    'Adds values when tickers are equal
                    totalOpen = totalOpen + ws.Cells(i, 3).Value
                    totalClose = totalClose + ws.Cells(i, 6).Value
                    yearlyChange = totalClose - totalOpen
                    'percentCh = (yearlyChange / totalOpen)
                    totalVol = totalVol + ws.Cells(i, 7).Value
            
                End If

    
            Next i
            
        Next ws

End Sub
