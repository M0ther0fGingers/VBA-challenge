Sub Tickers()

    ' Initiate variables
    Dim i As Integer
    Dim ws As Worksheet
  
    Set ws = ThisWorkbook.Sheets("A")
    
    
    
    ' Removed the work sheet loop for now until I get one page down
    ' For Each ws In Worksheets
    
        
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
        
        
        Opn = Cells(2, 3)
        
        'Label Column I
        Cells(1, 9).Value = "Ticker"
        
        'Label Column J
        Cells(1, 10).Value = "Change over Quarter"
        
        'Label Column K
        Cells(1, 11).Value = "Percent Change"
        
        'Label Column L
        Cells(1, 12).Value = "Stock Volume"
               
     
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
            Vol = (Vol + Cells(i, 7).Value)
           
    
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
           ' Find the closing value for the current ticker Tutor Sharon recommended change
            Cls = Cells(i, 6).Value
            
            ' Declare variable Ticker
             Ticker = Cells(i, 1).Value
                
            ' Calculate difference between Close and Open
            Change = Cls - Opn
            
            ' Calculate percentage, but not if 0
                    If Opn <> 0 Then
                        Range("K" & Summary_Table).Value = FormatPercent(Change / Opn, 2)
                        
                        Else
                            Range("K" & Summary_Table).Value = Null
                            
                    End If
                
            
            
            ' Add to the Ticker total
             '   Change = (Change + Cells(i, 8).Value) Tutor Sharon recommended change
                
          
                
                ' Print the Ticker name in the Summary Table
                  Range("I" & Summary_Table).Value = Ticker
                  
                ' Print the sum of changes in the Summary Table and format colors
                  Range("J" & Summary_Table).Value = Change
                    If Change > 0 Then
                        Range("J" & Summary_Table).Interior.ColorIndex = 4
                    Else
                        Range("J" & Summary_Table).Interior.ColorIndex = 3
                    End If
                
                  'Print the sum of Volume to the Summary Table
                  Range("L" & Summary_Table).Value = Vol
                  
                  ' Add one to the summary table row
                  Summary_Table = Summary_Table + 1
                  
                  ' Reset the Ticker Total
                  ' Change = 0
                  ' ChangeSummary = 0
                  Opn = Cells(i + 1, 3).Value
                  Vol = 0
                
            End If
            
        Next i
        
        Cells(2, 14).Value = "Greatest % Increase"
        Cells(3, 14).Value = "Greatest % Decrease"
        Cells(4, 14).Value = "Total Volume"
        
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
        
            If Cells(j, 11).Value > maxValue Then
                maxValue = Cells(j, 11).Value
                maxTicker = Cells(j, 9).Value
            End If
            
            If Cells(j, 11).Value < minValue Then
                minValue = Cells(j, 11).Value
                minTicker = Cells(j, 9).Value
            End If
            
            If Cells(j, 12).Value > maxVolume Then
                maxVolume = Cells(j, 12).Value
                maxVolTicker = Cells(j, 9).Value
            End If
        Next j
        Cells(2, 16).Value = FormatPercent(maxValue)
        Cells(3, 16).Value = FormatPercent(minValue)
        Cells(2, 15).Value = maxTicker
        Cells(3, 15).Value = minTicker
        Cells(4, 15).Value = maxVolTicker
        Cells(4, 16).NumberFormat = 0
        Cells(4, 16).Value = maxVolume
        
       'Cells(2, 14).Value = ws.Range("K2:K3")
        'Set Highest_Percent = Worksheets("A").Range("K2:K91")
        ' MsgBox (Highest_Percent)
        'answer = Application.WorksheetFunction.Max(Highest_Percent)
       ' MsgBox (Highest_Percent)
        
   ' Next ws
            



End Sub
