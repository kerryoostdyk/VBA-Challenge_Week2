Sub alphabetical_testing():

'assume stocks are pre-sorted
'Collect the ticker symbols using a loop

    'name some variables
    'maybe too many variables
    Dim ws As Worksheet
    Dim row As Long
    Dim row_count As Long
    Dim column As Integer
    Dim ticker As String
    Dim next_ticker As String
    Dim leaderboard_row As Integer
    Dim open_number As Double
    Dim closing_number As Double
    Dim quarter_change As Double
    Dim percent_change As Double
    Dim daily_change As Double
    Dim volume As LongLong
    Dim total_volume As LongLong
    Dim leaderboarder_row_count As Long

    
    For Each ws In ThisWorkbook.Worksheets
'Set title row
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Quarter Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
'Reset Quarterly change
    quarter_change = 0
    leaderboard_row = 2
    volume = 0
    total_volume = 0
    
'Name my  end row count
    row_count = Sheets("A").Cells(Rows.Count, "A").End(xlUp).row
    leaderboard_row_count = Cells(Rows.Count, "J").End(xlUp).row


        'Start the loop
            For row = 2 To row_count
            
            'extract the values from the workbook
            ticker = ws.Cells(row, 1).Value
            open_number = ws.Cells(2, 3).Value
            closing_number = ws.Cells(row, 6).Value
            next_ticker = ws.Cells(row + 1, 1).Value
            volume = ws.Cells(row, 7).Value
            
            'if statement
            If (ticker <> next_ticker) Then
            quarter_change = closing_number - open_number
            
            'percent change is quarter change / opening number
            total_volume = total_volume + volume
            percent_change = (quarter_change / open_number) * 100
            
            
            'put it on the leader board
            ws.Cells(leaderboard_row, 11).Value = quarter_change
            ws.Cells(leaderboard_row, 10).Value = ticker
            ws.Cells(leaderboard_row, 12).Value = percent_change
            ws.Cells(leaderboard_row, 13).Value = total_volume
            
            'Conditional Formating
            If (percent_change > 0) Then
                ws.Cells(leaderboard_row, 12).Interior.ColorIndex = 4
            ElseIf (percent_change < 0) Then
                ws.Cells(leaderboard_row, 12).Interior.ColorIndex = 3
            Else
                'Do Nothing
            
            End If
            
            'reset total
            quarter_change = 0
            leaderboard_row = leaderboard_row + 1
            total_volume = 0
            open_price = ws.Cells(row + 1, 3).Value
    
    Else
        'add total
        total_volume = total_volume + volume
        
    End If
    Next row
    

    'Second Loop for Second Leaderboard
    Dim greatest_per_inc As Double
    Dim greatest_per_dec As Double
    Dim greatest_total_vol As LongLong
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_total_ticker As String
    Dim j As Integer
    
    'init to first row of first leaderboard for comparison
    greatest_per_inc = ws.Cells(2, 12).Value
    greatest_per_dec = ws.Cells(2, 12).Value
    greatest_total_vol = ws.Cells(2, 13).Value
    greatest_increase_ticker = ws.Cells(2, 10).Value
    greatest_decrease_ticker = ws.Cells(2, 10).Value
    greatest_total_ticker = ws.Cells(2, 10).Value
        
    For j = 2 To leaderboard_row
        'compare current row to the first row
        If (ws.Cells(j, 12).Value > greatest_per_inc) Then
        'We have a new max
        greatest_per_inc = ws.Cells(j, 12).Value
        greatest_increase_ticker = ws.Cells(j, 10).Value

    End If
    
        
    'put it on the leaderboard part 2
    ws.Cells(2, 17).Value = greatest_per_increase
    ws.Cells(3, 17).Value = greatest_per_decrease
    ws.Cells(4, 17).Value = greatest_total_vol
    greatest_increase_ticker = ws.Cells(2, 16).Value
    greatest_decrease_ticker = ws.Cells(2, 16).Value
    greatest_total_ticker = ws.Cells(2, 16).Value
    
    
   'repeat for the next few greatest
    Next j
    
    ws.Range("P2").Value = greatest_increase_ticker
    ws.Range("P3").Value = greatest_decrease_ticker
    ws.Range("P4").Value = greatest_total_ticker
    
    Next ws
    
End Sub

Sub reset()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Delete columns I through P
        ws.Range("I:P").Delete
    Next ws










