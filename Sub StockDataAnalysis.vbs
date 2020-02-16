Sub StockDataAnalysis()

    Dim sht As Worksheet
    Dim lastRow As Double
    Dim ticker As String
    Dim FirstDayOpen As Single
    Dim s As Long
    Dim TotalStockVolume As Double
    
    For Each ws In Worksheets
        ws.Select
    
        ws.Range("I:Q").EntireColumn.Clear
    
    
        'Test that the sheets are looping through
        'MsgBox "The current sheet is: " & ws.Name
        
        'Test that we are picking up the last row in each sheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'MsgBox "The last row on this sheet is: " & lastRow
        
        'Create headings for the summary table
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Price Change "
        Cells(1, 11).Value = "Price Change Percentage"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Add headings for Greatest Percentage Increase, Decrease and Volumes
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
       
        
        'First set the ticker equal to A2
        ticker = Cells(2, 1).Value
        
        'Set the FirstDayOpen value equal to C2
        FirstDayOpen = Cells(2, 3).Value
        
                'For the current sheet, let's process all the rows
                For i = 2 To lastRow
                
                    'If the the ticker in the current row matches the value of the ticker variable then
                    'Just keep adding to the total volume for that ticker
                    
                    If Cells(i, 1).Value = ticker Then
                    
                        TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                    
                    Else 'it means the ticker no longer matches so we have actually passed the last row for the ticker
                        
                        'set the value for LastDayClose
                        LastDayClose = Cells(i - 1, 6).Value
                        s = s + 1
                        
                        'Now we have all the values we need to write to the table
                        'Write Values to the summary table
                        Cells(1 + s, 9).Value = ticker
                        Cells(1 + s, 10).Value = (LastDayClose - FirstDayOpen)
                        Cells(1 + s, 10).NumberFormat = "0.00"
                        
                        If TotalStockVolume = 0 Then
                        
                            Cells(1 + s, 11).Value = 0
                            
                        ElseIf TotalStockVolume > 0 And FirstDayOpen = 0 Then
                        
                            Cells(1 + s, 11).Value = (LastDayClose - FirstDayOpen) / 1
                            
                        Else
                        
                            Cells(1 + s, 11).Value = (LastDayClose - FirstDayOpen) / FirstDayOpen
                            
                        End If
                        
                        Cells(1 + s, 11).NumberFormat = "0.00%"
                        
                        Cells(1 + s, 12).Value = TotalStockVolume
                    
                        'Apply color coding
                        'We want to make the color of the cell in the Yearly Change column red if the yearly change is negative _
                         we want to make the color of the cell in the Yearly Change column green if the yearly change is positive
                        
                        'Perform the test and conditionally apply color
                        If Cells(1 + s, 10).Value > 0 Then
                           Cells(1 + s, 10).Interior.ColorIndex = 4
                            
                        ElseIf Cells(1 + s, 10).Value < 0 Then
                                 Cells(1 + s, 10).Interior.ColorIndex = 3
                            
                        End If
                        
                        
                        'set the new values for ticker, FirstDayOpen and Total Stock Volume
                        ticker = Cells(i, 1).Value
                        TotalStockVolume = Cells(i, 7).Value
                        FirstDayOpen = Cells(i, 3).Value
                        
                    End If
                    
                    
            
                Next i
        
        
        'Sort to get the Largest % Increase
        Range(Cells(1, 9), Cells(s + 1, 12)).Sort Cells(1, 11), xlDescending, , , , , , xlYes
        
        Cells(2, 16).Value = Cells(2, 9).Value
        Cells(2, 17).Value = Cells(2, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
        
        
        'Sort to get the Lowest % Increase
        Range(Cells(1, 9), Cells(s + 1, 12)).Sort Cells(1, 11), , , , , , , xlYes
        Cells(3, 16).Value = Cells(2, 9).Value
        Cells(3, 17).Value = Cells(2, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
        
        'Sort to get the highest Total Stock Volume
        Range(Cells(1, 9), Cells(s + 1, 12)).Sort Cells(1, 12), xlDescending, , , , , , xlYes
        Cells(4, 16).Value = Cells(2, 9).Value
        Cells(4, 17).Value = Cells(2, 12).Value
        Cells(4, 17).NumberFormat = "0.00"
       
        'sort summary data by ticker in ascending order
        Range(Cells(1, 9), Cells(s + 1, 12)).Sort Cells(1, 9), xlAscending, , , , , , xlYes
       
       
    'reset i and s
    i = 2
    s = 0
     'Now you are done so move on to the next sheet
     
    Next ws
    
    
End Sub

