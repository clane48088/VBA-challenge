Attribute VB_Name = "Module1"
 Sub stockAnalysis():
  
   
    Dim total As Double 'total stock volumn
    Dim row As Long 'loop control variablethat will go through the rows of the sheet
    Dim rowCount As Long 'variable will hold the number of rows in a sheet
    Dim quaterlychange As Double  'variable that hold the quarterly change for each stock in a sheet
    Dim percentChange As Double ' varibable that holds the percent change for each stock in a sheet
    Dim summaryTableRow As Long ' variable holds the rows of the summary table row
    Dim stockStartRow As Long ' variable that holds the start pf a stock's rows in the sheet
    Dim startValue As Long  ' start row for a stock (location of first open)
    Dim lastTicker As String  ' finds the last ticker in the sheet
            
    For Each ws In Worksheets
            
        ' loop through all of the worksheets
        
             ' add a title
             ws.Range("I1").Value = "Ticker"
             ws.Range("J1").Value = "Quarterly Change"
             ws.Range("K1").Value = "Percent Change"
             ws.Range("L1").Value = "Total Stock Value"
             ws.Range("P1").Value = "Ticker"
             ws.Range("Q1").Value = "Value"
             ws.Range("O2").Value = "Greatest % Increase"
             ws.Range("O3").Value = "Greatest % Decrease"
             ws.Range("O4").Value = "Greatest Total Volumn"
             
             ' initialize the values
             summaryTableRow = 0 ' summary table row starts at 0 in the sheet
             total = 0 ' total stock starts at 0
             quarterlyChange = 0 ' quarterly change starts at 0
             stockStartRow = 2 ' first stock in the sheet is starts at row 2
             startValue = 2 ' first open of the first stock value is on row 2
             
             ' get the value of the last row
             rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
             
             ' find the last ticker to get out of the loop
             lastTicker = ws.Cells(rowCount, 1).Value
             
             ' loops to the end of the sheet
             For row = 2 To rowCount
             
                'check to see if there are changes in column A (1st column)
                If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                
                   'Calculate the total one last time for the ticker
                   total = total + ws.Cells(row, 7).Value ' Grabs the value from the 7th colun which is (G)
                   
                   ' check to see if the value of the total volume is 0
                   If total = 0 Then
                    ' print the results (columns, I, J, K, and L)
                     ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value ' prints the stock name in column I
                     ws.Range("J" & 2 + summaryTableRow).Value = 0 ' prints 0 in column J
                     ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%" 'prints 0% in colmun K (% change)
                     ws.Range("L" & 2 + summaryTableRow).Value = 0 ' prints 0 in colmun L (total stock volumn)
                   Else
                    ' find the first non-zero starting value
                    If ws.Cells(startValue, 3).Value = 0 Then
                        For findValue = stockStartRow To row
                            
                            ' check to see if the next (or next) value does not equal 0
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                startValue = findValue
                                ' once we have a non-zero value, break out the loop
                                Exit For
                            End If
                        Next findValue
                    End If
                        
                        ' quarterly change (difference in last close - first open)
                        quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
                        
                        ' the percent change (yearly change / first open)
                        percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
                            
                        ' print the results (columns, I, J, K, and L)
                        ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value ' prints the stock name in column I
                        ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange ' prints in column J (quarter change)
                        ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00" ' formats column J
                        ws.Range("K" & 2 + summaryTableRow).Value = percentChange 'prints in colmun K (% change)
                        ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%" ' formats colmun K
                        ws.Range("L" & 2 + summaryTableRow).Value = total ' prints 0 in colmun L (total stock volumn)
                        ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###" ' formats colmun l
                        
                        ' formatting for yearly change coloumn
                        If quarterlyChange > 0 Then
                            
                            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 ' green = +
                        ElseIf quarterlyChange < 0 Then
                            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3 ' red = -
                        Else
                            ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0 ' white/no change
                                        
                         
                    End If
                   
                    'resets the totals
                    total = 0
                    ' quarterly change reset
                    averageChange = 0   ' resets the average change for the next ticker
                    quarterlyChange = 0
                    startValue = row + 1    ' moves the start row to the next row in the sheet
                    ' move to the next row in the summary table
                    summaryTableRow = summaryTableRow + 1
                    
                End If
            
            Else
                ' if the ticker is the same
                total = total + ws.Cells(row, 7).Value ' Grabs the value from the 7th column which is (G)
                    
            End If
                
                     
         Next row
            
         ' clean up
         ' find the last row of data in the summary table by finding the last ticker in the summary section
           
         ' update the summary table row
           summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
            
          ' last data in the extra rows from columns J-L
          Dim lastExtraRow As Long
          lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
                    
                            
          ' loop that clears the extra data from columns I - L
          For e = summaryTableRow To lastExtraRow
              ' for loopthat goes through columns I-L (9-12)
              For Column = 9 To 12
                 ws.Cells(e, Column).Value = ""
                 ws.Cells(e, Column).Interior.ColorIndex = 0
              Next Column
              
          Next e
          
          ' print the summary aggregates
          ' after generating the info in the summary section, find the greates % increase and decrease, then find the greatest total volumn
          ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
          ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
          ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
          
          ' use Match () to find the row numbers of the ticker names associated with the greatest % increase and decrease, then find the greates total stock vol
          Dim greatestIncreaseRow As Double
          Dim greatestDecreaseRow As Double
          Dim greatestTotVolRow As Double
          greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
          greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
          greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
          
          ' display the ticker - greatest increase, greatest decrease and greatest total stock volumn
          ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
          ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
          ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
        
         'format the summary table columns
         For s = 0 To summaryTableRow
            ws.Range("J" & 2 + s).NumberFormat = "0.00"       ' Formats Quarterly Change
            ws.Range("K" & 2 + s).NumberFormat = "0.00%"       ' Formats Percent Changes
            ws.Range("L" & 2 + s).NumberFormat = "#,###"      ' formats the total stock volumn
            
         Next s
         
         ' format the summary aggregates for the worksheets
         ws.Range("Q2").NumberFormat = "0.00%"     'format the greatest & increase
         ws.Range("Q3").NumberFormat = "0.00%"     'format the greatest & decrease
         ws.Range("Q4").NumberFormat = "#,###"     'format the greatest total stock volumn
         
          ' Autofit the columns for all worksheets
          ws.Columns("A:Q").AutoFit
          
     Next ws
     
              
     
    
End Sub
