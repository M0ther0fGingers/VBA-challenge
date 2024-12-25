Sub Tickers()

    ' Initiate variables
    Dim i As Integer
    Dim ws As Worksheet
  
    'Set ws = ThisWorkbook.Sheets("A")

    'Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
 
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
        
        ' Set the initial Change value
        ' Change = 0 Tutor Sharon recommended change
        
        ' Declare variable Volumn
          Vol = 0
        
        ' Loop through the rows
        For i = 2 To LastRow
    
           'Add to the Volume total for each stock
            Vol = (Vol + ws.Cells(i, 7).Value)
           
    
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           ' Find the closing value for the current ticker Tutor Sharon recommended change
            Cls = ws.Cells(i, 6).Value
            
            ' Declare variable Ticker
             Ticker = ws.Cells(i, 1).Value
                
            ' Calculate difference between Close and Open
            Change = Cls - Opn
            
            ' Calculate percentage, but not if 0
                    If Opn <> 0 Then
                        ws.Range("K" & Summary_Table).Value = FormatPercent(Change / Opn, 2)
                        
                        Else
                            ws.Range("K" & Summary_Table).Value = Null
                            
                    End If
                
            
            
            ' Add to the Ticker total
             '   Change = (Change + Cells(i, 8).Value) Tutor Sharon recommended change
                
          
                
                ' Print the Ticker name in the Summary Table
                  ws.Range("I" & Summary_Table).Value = Ticker
                  
                ' Print the sum of changes in the Summary Table and format colors
                  ws.Range("J" & Summary_Table).Value = Change
                    If Change > 0 Then
                        ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
                    End If
                
                  'Print the sum of Volume to the Summary Table
                  ws.Range("L" & Summary_Table).Value = Vol
                  
                  ' Add one to the summary table row
                  Summary_Table = Summary_Table + 1
                  
                  ' Reset the Ticker Total
                  ' Change = 0
                  ' ChangeSummary = 0
                  Opn = ws.Cells(i + 1, 3).Value
                  Vol = 0
                
            End If
            
        Next i
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Total Volume"
        
        Dim maxValue As Double
        Dim minValue As Double
        Dim maxTicker As String
        Dim minTicker As String
        Dim maxVolume As LongLong
        Dim maxVolTicker As String
        
        
        
        maxValue = 0
        minValue = 100
        maxVolume = 0
        
        
        For j = 2 To LastRow
        
            If ws.Cells(j, 11).Value > maxValue Then
                maxValue = ws.Cells(j, 11).Value
                maxTicker = ws.Cells(j, 9).Value
            End If
            
            If ws.Cells(j, 11).Value < minValue Then
                minValue = ws.Cells(j, 11).Value
                minTicker = ws.Cells(j, 9).Value
            End If
            
            If ws.Cells(j, 12).Value > maxVolume Then
                maxVolume = ws.Cells(j, 12).Value
                maxVolTicker = ws.Cells(j, 9).Value
            End If
        Next j
        ws.Cells(2, 16).Value = FormatPercent(maxValue)
        ws.Cells(3, 16).Value = FormatPercent(minValue)
        ws.Cells(2, 15).Value = maxTicker
        ws.Cells(3, 15).Value = minTicker
        ws.Cells(4, 15).Value = maxVolTicker
        ws.Cells(4, 16).NumberFormat = 0
        ws.Cells(4, 16).Value = maxVolume

   Next ws

End Sub
