Sub StockPrice2018():

'Declaring & Assigning the variables
           'This is to work on each worksheet at the same time
            Dim ws As Worksheet
            For Each ws In Worksheets
            Dim TickerName As String
            Dim OpeningPrice, ClosingPrice, YearlyChange As Double
        
        Dim Summary_Table_Row As Integer
        'Assign Value to the Variable
        Summary_Table_Row = 2
        Dim TotalVolume As Double
        'Assign Value to the Variable
        TotalVolume = 0
        'Define Last Row in each row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim YearBegin, YearEnd As Long
       'Assign the First value of the Open Price Column to the Variable
       OpeningPrice = ws.Cells(2, 3).Value
       'For Summary Table one Heading
          '-----------For the second Summary Table
           ' Summary Table1 Headings
            ws.Range("J1").Value = "Ticker"
            ws.Range("J1").Font.Bold = True
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("K1").Font.Bold = True
            ws.Range("L1").Value = "Percent Change"
            ws.Range("L1").Font.Bold = True
            ws.Range("M1").Value = "Total Stock Volume"
            ws.Range("m1").Font.Bold = True
            
 For i = 2 To LastRow

     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the Ticker Name and place it in to column J
             TickerName = ws.Cells(i, 1).Value
             
             ws.Range("J" & Summary_Table_Row).Value = TickerName
                 
                 ' Calculate the Yearly Change and Coeersponding Percentage Change using below calculation and place it in to column K & L
                         ClosingPrice = ws.Cells(i, 6).Value
                         
                         YearlyChange = ClosingPrice - OpeningPrice
                         
                         PercentageChange = YearlyChange / OpeningPrice
                         
                        ws.Range("K" & Summary_Table_Row).Value = YearlyChange
                        ws.Range("L" & Summary_Table_Row).Value = PercentageChange
                        'Formatting to % Form
                        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                  'Add to the TotalVolume
                     TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                  'Print the TotalVolume to the Summary Table Total Stock Volume Column
                     ws.Range("M" & Summary_Table_Row).Value = TotalVolume
                  'Changing cellcolor depends upon the +ve/-Ve Values
                  
                        If ws.Range("K" & Summary_Table_Row).Value > 0 Then
                           ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                        ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
                           ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                        End If
       
                 'Reset the Variables
    
                 Summary_Table_Row = Summary_Table_Row + 1
                 OpeningPrice = ws.Cells(i + 1, 3).Value
                 'Reset the Brand Total
                 TotalVolume = 0
            
                ' If the cell immediately following a row is the same brand...
     Else

            ' Add to the Brand Total
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value

     End If
Next i
    
'-----------For the second Summary Table
' Summary Table1 Headings
            ws.Range("P1").Value = "Summary Table"
            ws.Range("P1").Font.Bold = True
            ws.Range("P3").Value = "Greatest% Increase"
            ws.Range("P4").Value = "Greatest% Decrease"
            ws.Range("P5").Value = "Greatest Total Volume"
            ws.Range("Q2").Value = "Ticker"
            ws.Range("R2").Value = "Value"

      'Find the Maximum % increse and corresponding Ticker Value
            ws.Range("R3").Value = WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("R3").NumberFormat = "0.00%"
            max_row_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
            ws.Range("Q3").Value = ws.Cells(max_row_number + 1, 10)
            
      'Find the Minimum % increse and corresponding Ticker Value
            ws.Range("R4").Value = WorksheetFunction.Min(ws.Range("L:L"))
            ws.Range("R4").NumberFormat = "0.00%"
            Min_row_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
            ws.Range("Q4").Value = ws.Cells(Min_row_number + 1, 10)
            
      
      'Find the Maximum value of Total Volume Column and corresponding Ticker Value
            ws.Range("R5").Value = WorksheetFunction.Max(ws.Range("M:M"))
            Greatest_Total = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & Summary_Table_Row)), ws.Range("M2:M" & Summary_Table_Row), 0)
            ws.Range("Q5").Value = ws.Cells(Greatest_Total + 1, 10)
                
      
 Next ws
 End Sub
