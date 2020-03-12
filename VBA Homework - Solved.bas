Attribute VB_Name = "Module1"
'Option Explicit
'Create a script that will loop through all the stocks for one year and
'output the following information:
'-----The ticker symbol.
'-----Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'-----The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'-----The total stock volume of the stock.
'You should also have conditional formatting that will:
' -----highlight positive change in green and
'-----highlight negative change in red.
'--------------------------------START OF PROGRAM-----------------------------

Sub Stocks()
    '-----Loop through all the worksheets-----
    For Each x In Worksheets
        'Calculate Last Row
        LastRow = x.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Define variables for looping
        Dim i As Long
              
        'Define variables for identifying current and next ticker
        Dim CurrentTicker As String
        Dim NextTicker As String
        
        'Define variables for summary table calculations
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim CurrentVolume As Long
        Dim SummaryRow As Long
                
        'Assign initial values to all calculated variables
        SummaryRow = 2
        OpeningPrice = x.Range("C2").Value
        YearlyChange = 0
        PercentChange = 0
        TotalVolume = 0
        
       'Assign Headers for the first summary table
        x.Range("I1").Value = "Ticker"
        x.Range("J1").Value = "Yearly Change"
        x.Range("K1").Value = "% Change"
        x.Range("L1").Value = "Total Stock Volume"
        
         'Assign Headers for the second summary table
        x.Range("N2").Value = "Greatest % Increase"
        x.Range("N3").Value = "Greatest % Decrease"
        x.Range("N4").Value = "Greatest Total Volume"
        x.Range("O1").Value = "Ticker"
        x.Range("P1").Value = "Value"
        
        For i = 2 To LastRow
        'Assign the value of the current and next tickers
        CurrentTicker = x.Cells(i, 1).Value
        NextTicker = x.Cells(i + 1, 1).Value
        
        If CurrentTicker <> NextTicker Then
            'Print ticker in summary table
            x.Cells(SummaryRow, 9).Value = CurrentTicker
            
            'Find Closing price for ticker
            ClosingPrice = x.Cells(i, 6).Value
            
            'Calculate YearlyChange and store in summary table
            YearlyChange = ClosingPrice - OpeningPrice
            x.Cells(SummaryRow, 10).Value = YearlyChange
            ' x.Cells(SummaryRow, 10).Style = "Standard"
            
            'Perform Conditional formatting for YearlyChange in summary table
            If YearlyChange > 0 Then
                x.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
                x.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            End If
           
            'If OpeningPrice=0 then assign YearlyChange=0
            If OpeningPrice = 0 Then
                PercentChange = 0
            Else
            'Calculate PercentChange and store in summary table
             PercentChange = (YearlyChange / OpeningPrice)
             x.Cells(SummaryRow, 11).Value = PercentChange
             x.Cells(SummaryRow, 11).Style = "Percent"
            End If
            
            'Calculate total volume
            CurrentVolume = x.Cells(i, 7).Value
            TotalVolume = (TotalVolume + CurrentVolume)
            x.Cells(SummaryRow, 12).Value = TotalVolume
                       
           'Assign value of Opening Price for next ticker
            OpeningPrice = x.Cells(i + 1, 3).Value
            
            'Go to next row in summary table
            SummaryRow = SummaryRow + 1
            
            'Initialize TotalVolume for next ticker
            TotalVolume = 0
        Else
            'Calculate total volume
            CurrentVolume = x.Cells(i, 7).Value
            TotalVolume = (TotalVolume + CurrentVolume)
            x.Cells(SummaryRow, 12).Value = TotalVolume
        End If
        Next i
        
        '-------Start of Challenge - Populate the second Summary table ---------
         Dim j As Integer
         
        'Calculate Last or of first summary table
        LastSumRow = x.Cells(Rows.Count, 11).End(xlUp).Row
        
        'Initialize Maximum Percent Increase
         Dim MaxIncrease As Double
         Dim MaxIncTicker As String
         MaxIncrease = x.Cells(2, 11).Value
                  
        'Initialize Maximum Percent Decrease
        Dim MaxDecrease As Double
        Dim MaxDecTicker As String
        MaxDecrease = x.Cells(2, 11).Value
        
        'Initialize Maximum Volume
        Dim MaxVolTicker As String
        MaxVolume = x.Cells(2, 12).Value
        
        'Loop through the values in first Summary table to calculate the max values
        For j = 3 To LastSumRow
            
            If x.Cells(j, 11).Value > MaxIncrease Then
                MaxIncrease = x.Cells(j, 11)
                MaxIncTicker = x.Cells(j, 9).Value
                x.Range("O2").Value = MaxIncTicker
                x.Range("P2").Value = MaxIncrease
                x.Range("P2").Style = "Percent"
            End If
            
            If x.Cells(j, 11).Value < MaxDecrease Then
                MaxDecrease = x.Cells(j, 11)
                MaxDecTicker = x.Cells(j, 9).Value
                x.Range("O3").Value = MaxDecTicker
                x.Range("P3").Value = MaxDecrease
                x.Range("P3").Style = "Percent"
            End If
            
            If x.Cells(j, 12).Value > MaxVolume Then
                MaxVolume = x.Cells(j, 12).Value
                MaxVolTicker = x.Cells(j, 9).Value
                x.Range("O4").Value = MaxVolTicker
                x.Range("P4").Value = MaxVolume
            End If
            
        Next j
        '---------End of Challenge - Populate second summary table---------
        
    Next x

End Sub
'--------------------------------END OF PROGRAM-----------------------------



