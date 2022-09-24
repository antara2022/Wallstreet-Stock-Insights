Attribute VB_Name = "Module1"
' Steps:
' ----------------------------------------------------------------------------

' Create a script that loops through all the stocks for one year and outputs the following information:
' 1. The ticker symbol
' 2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' 3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' 4. The total stock volume of the stock.

Sub MultipleYearStockData():

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        ' Define everything
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        Dim TickerCount As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        Dim PercentChange As Double
        Dim GreatestIncr As Double
        Dim GreatestDecr As Double
        Dim GreatestVol As Double
        
        ' Get WorksheetName
        WorksheetName = ws.Name
        
        ' Create headers for columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Setting Ticker Counter to first row
        TickerCount = 2
        
        ' Set starting row
        j = 2
        
        ' Find last non-empty cell in column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in Column A is " & LastRowA)
        
            ' Looping through all rows
            For i = 2 To LastRowA
            
                ' Check to see if ticker name is different
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Write ticker in column I
                ws.Cells(TickerCount, 9).Value = ws.Cells(i, 1).Value
                
                ' Add Yearly Change to Column J
                ws.Cells(TickerCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                ' Conditional formatting
                If ws.Cells(TickerCount, 10).Value < 0 Then
                
                ' Set cell background color to red to highlight negative change
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
                
                Else
                
                ' Set cell background color to green to highlight positive change
                ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
                
                End If
                
                ' Add Percent Change to Column K
                If ws.Cells(j, 3).Value <> 0 Then
                PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                ' Format to Percent Sign
                ws.Cells(TickerCount, 11).Value = Format(PercentChange, "Percent")
                
                Else
                
                ws.Cells(TickerCount, 11).Value = Format(0, "Percent")
                
                End If
                
            ' Add Total Stock Volume into Column L
            ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            ' Increase TickerCount by 1
            TickerCount = TickerCount + 1
            
            ' Set new start row for ticker block
            j = i + 1
            
            End If
            
        Next i
        
    ' Find last non-empty cell in Column I
    LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'MsgBox ("Last row in column I is " & LastRowI)
    
    ' Create summary
    GreatestVol = ws.Cells(2, 12).Value
    GreatestIncr = ws.Cells(2, 11).Value
    GreatestDecr = ws.Cells(2, 11).Value
    
        ' Create a loop for summary
        For i = 2 To LastRowI
        
        ' Greatest Total Volume
        If ws.Cells(i, 12).Value > GreatestVol Then
        GreatestVol = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        GreatestVol = GreatestVol
        
        End If
        
        ' Greatest increase
        If ws.Cells(i, 11).Value > GreatestIncr Then
        GreatestIncr = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        GreatestIncr = GreatestIncr
        
        End If
        
        ' Greatest decrease
        If ws.Cells(i, 11).Value < GreatestDecr Then
        GreatestDecr = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        
        Else
        
        GreatestDecr = GreatestDecr
        
        End If
        
        'Write summary results
        ws.Cells(2, 17).Value = Format(GreatestIncr, "Percent")
        ws.Cells(3, 17).Value = Format(GreatestDecr, "Percent")
        ws.Cells(4, 17).Value = Format(GreatestVol, "Scientific")
        
        Next i
        
        ' Adjust column width
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
       
End Sub
