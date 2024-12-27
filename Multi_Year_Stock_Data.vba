Sub Tickers_Analysis()

    'Initiate variables
    Dim i As Long
    Dim ws As Worksheet

    'Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
    ' Declare variables
        Dim LastRow As Long
        Dim Ticker As String
        Dim Opn As Double
        Dim Cls As Double
        Dim Vol As Double
        Dim Summary_Table As Integer
        Dim Ticker_Total As Double
        Dim Change As Double
        Dim ChangeSummary As Double
        Dim Percent_Change As Double
        Dim Highest_Percent As Double
        Dim Lowest_Percent As Double
        
        'Starting value for variable Opn
        Opn = ws.Cells(2, 3)

        'Label Column I
        ws.Cells(1, 9).Value = "Ticker"
        
        'Label Column J
        ws.Cells(1, 10).Value = "Change over Quarter"
        
        'Label Column K
        ws.Cells(1, 11).Value = "Percent Change"
        
        'Label Column L
        ws.Cells(1, 12).Value = "Stock Volume"
               
        ' calculate the last row in each sheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
      
        ' Indicate where the data should print
        Summary_Table = 2
    
        ' Declare variable Volumn
        Vol = 0
        
        ' Loop through the rows
         For i = 2 To LastRow
        
            'Add to the Volume total for each stock
            Vol = (Vol + ws.Cells(i, 7).Value)
         
            'Compare the value of cells in column A. If the value is not equal, then move on to next line.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Find the closing value for the current ticker
                Cls = ws.Cells(i, 6).Value
        
                'Declare variable Ticker
                Ticker = ws.Cells(i, 1).Value
            
                'Calculate difference between Close and Open
                Change = Cls - Opn
        
            'Calculate percentage, but not if 0
                If Opn <> 0 Then
                        ws.Range("K" & Summary_Table).Value = FormatPercent(Change / Opn, 2)
                        
                        'Print the percent change to the summary table
                        Else
                            ws.Range("K" & Summary_Table).Value = Null
                End If
                
            ' Print the Ticker name in the Summary Table
              ws.Range("I" & Summary_Table).Value = Ticker
              
            ' Print the sum of changes in the Summary Table
              ws.Range("J" & Summary_Table).Value = Change
                  
                  'Format summary table colors
                    If Change > 0 Then
                        ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
                    End If
                
                  'Print the sum of Volume to the Summary Table
                  ws.Range("L" & Summary_Table).Value = Vol
                  
                  ' Add one to the summary table row
                  Summary_Table = Summary_Table + 1
                  
                  ' Reset values before next loop starts
                  Opn = ws.Cells(i + 1, 3).Value
                  Vol = 0
            End If
            
        Next i
        
        'Label outliers summary
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Total Volume"
        
        'Declare values for outliers summary
        Dim maxValue As Double
        Dim minValue As Double
        Dim maxTicker As String
        Dim minTicker As String
        Dim maxVolume As LongLong
        Dim maxVolTicker As String
        
        'Starting values for outliers summary
        maxValue = 0
        minValue = 100
        maxVolume = 0
        
        'Loop through column J to find outliers
        For j = 2 To LastRow
        
            'Find the max percentage
            If ws.Cells(j, 11).Value > maxValue Then
                maxValue = ws.Cells(j, 11).Value
                maxTicker = ws.Cells(j, 9).Value
            End If
            
            'Find the min percentage
            If ws.Cells(j, 11).Value < minValue Then
                minValue = ws.Cells(j, 11).Value
                minTicker = ws.Cells(j, 9).Value
            End If
            
            'Find the max volumn
            If ws.Cells(j, 12).Value > maxVolume Then
                maxVolume = ws.Cells(j, 12).Value
                maxVolTicker = ws.Cells(j, 9).Value
            End If
            
        Next j
        
        'Format outliers values
        ws.Cells(2, 16).Value = FormatPercent(maxValue)
        ws.Cells(3, 16).Value = FormatPercent(minValue)
        ws.Cells(2, 15).Value = maxTicker
        ws.Cells(3, 15).Value = minTicker
        ws.Cells(4, 15).Value = maxVolTicker
        ws.Cells(4, 16).NumberFormat = 0
        ws.Cells(4, 16).Value = maxVolume

   Next ws

End Sub



