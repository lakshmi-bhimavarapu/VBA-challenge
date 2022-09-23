Attribute VB_Name = "Module1"
Sub GetUnique()

Dim count, j As Integer

'Counter to store the total number of sheet in the workbook
count = Application.Worksheets.count

    For j = 1 To count
        'Activate one sheet at a time and do the processing
        Worksheets(j).Activate

        Dim u, lastrow, finalrow, Bottom As Long
        Dim i, presentvalue As Variant

        'Set the header for the new rows
        Cells(1, 8).Value = "Ticker"
        Cells(1, 9).Value = "Yearly Change"
        Cells(1, 10).Value = "Percent Change"
        Cells(1, 11).Value = "Total Stock Volume"

i = 2

'A variable to store the last non-blank row number
lastrow = Cells(Rows.count, 1).End(xlUp).Row


    For i = 2 To lastrow
    'Variable u holds the first ticker in column 1
    u = Cells(i, 1).Value

    'Variable to hold the bottom row number for that particular ticker
    Bottom = Range("A:A").Find(what:=u, LookAt:=xlWhole, MatchCase:=True, searchdirection:=xlPrevious).Row
    
    'Increment the iteration to start on new ticker using above variable
    i = Bottom + 1

    'Variable to hold the last row of unique tickers
    finalrow = Cells(Rows.count, 8).End(xlUp).Row
        
    Range("H2").Select
    
    'Variable to hold the ticker among the unique tickers written to new column
    Set presentvalue = Range("H2:H" & finalrow).Find(what:=u, LookAt:=xlWhole)
    
    If presentvalue Is Nothing Then
    Cells(finalrow, 8).Offset(1, 0).Value = u
    Else
    
    End If
    
    Next i
    
    Next j

'Call new sub to calculate the price changes
Call PriceChange

End Sub


Sub PriceChange()

Dim count, k As Integer

'Counter to store the total number of sheet in the workbook
count = Application.Worksheets.count

    'Counter to store the total number of sheet in the workbook
    For k = 1 To count
        Worksheets(k).Activate

Dim year_begin_open, year_end_close, stock_volume, highstock As Double
Dim u As String
Dim Top, Bottom, i, j, highest, lowest, p, LRN, HRN, SRN As Long
Dim MatchingArray() As Variant
 
    'Setting the header for new columns
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"

'Finding the row number of the last unique ticker
lastrow = Cells(Rows.count, 8).End(xlUp).Row

'Finding the row number of the last row of the whole dataset
finalrow = Cells(Rows.count, 1).End(xlUp).Row

For i = 2 To lastrow
    u = Cells(i, 8).Value
    
    'Get the opening price for the ticker
    Top = Range("A:A").Find(what:=u, LookAt:=xlWhole, MatchCase:=True).Row
    year_begin_open = Cells(Top, 3).Value
    
    'Get the closing price for the ticker by searching from the bottom of the file
    Bottom = Range("A:A").Find(what:=u, LookAt:=xlWhole, MatchCase:=True, searchdirection:=xlPrevious).Row
    year_end_close = Cells(Bottom, 6).Value
    
    'Calculating the Price Change
    Cells(i, 9).Value = year_end_close - year_begin_open
    
    'Coloring the Price Change based on positive or negative change
    If Cells(i, 9).Value > 0 Then
        Cells(i, 9).Interior.ColorIndex = 4
    ElseIf Cells(i, 9).Value < 0 Then
        Cells(i, 9).Interior.ColorIndex = 3
    Else: Cells(i, 9).Interior.ColorIndex = xlNone
    End If
    
    'Calculating the Percent Change
    Cells(i, 10).Value = Cells(i, 9).Value / year_begin_open
    
    stock_volume = 0
    
    'Calculating the stock volume for each ticker
    For j = Top To Bottom
    stock_volume = Cells(j, 7).Value + stock_volume
    Next j
    
    Cells(i, 11).Value = stock_volume
    
Next i
    
    'Check for "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
    
    lowest = Application.WorksheetFunction.Min(Range("J:J"))
    highest = Application.WorksheetFunction.Max(Range("J:J"))
    highstock = Application.WorksheetFunction.Max(Range("K:K"))
   
    MatchingArray = Range("A:K")
    
    For p = 2 To finalrow
    
    'Find the lowest row number(LRN) using lowest value that we calculated above
    If StrComp(MatchingArray(p, 10), lowest, vbTextCompare) = 0 Then
          LRN = p
         Exit For
     End If
    Next p
    
    Cells(3, 15).Value = Cells(LRN, 8).Value
    Cells(3, 16).Value = Cells(LRN, 10).Value
    
    For p = 2 To finalrow
    
    'Find the highest row number(HRN) using highest value that we calculated above
    If StrComp(MatchingArray(p, 10), highest, vbTextCompare) = 0 Then
          HRN = p
         Exit For
     End If
    Next p
    
    Cells(2, 15).Value = Cells(HRN, 8).Value
    Cells(2, 16).Value = Cells(HRN, 10).Value
    
    For p = 2 To finalrow
    
    'Find the stock row number(SRN) using highest stock value that we calculated above
    If StrComp(MatchingArray(p, 11), highstock, vbTextCompare) = 0 Then
         SRN = p
         Exit For
     End If
    Next p
    
    Cells(4, 15).Value = Cells(SRN, 8).Value
    Cells(4, 16).Value = Cells(SRN, 11).Value
       
    Range("J:J").NumberFormat = "0.00%"
    Range("P2:P3").NumberFormat = "0.00%"
    Range("A:P").Columns.AutoFit
    
    'In order to match my output to what is provided in "hard solution" screenprint, I inserted an empty column
    Range("H:H").EntireColumn.Insert
    
    Next k

End Sub
