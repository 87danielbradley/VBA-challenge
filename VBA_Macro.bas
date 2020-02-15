Attribute VB_Name = "Module2"
Sub Master()
'I used this source to get the code to run on multiple sheets
'<https://www.ozgrid.com/forum/index.php?thread/107423-run-the-same-macro-on-multiple-sheets-in-same-workbook/>
'The rest was written myself

Dim AllSheets As Worksheet
For Each AllSheets In ThisWorkbook.Worksheets
AllSheets.Select




'FirstOpen holds the Opening ticker value
Dim FirstOpen As Variant
'LastClose holds the value for the Final Stock Close
Dim LastClose As Variant
'Yearly Change = LastClose - First Open
Dim YearlyChange As Variant


'count is used to reorder Unique tickers
Dim Count As Integer
Count = 2

'CellNameTemp is a temporary variable for ticker label to find Unique
Dim CellNameTemp As String
'I defined it to ensure the first statement would be false initially.  It no longer matters as it gets redefined immediately
CellNameTemp = "empty"

'This variable is being used to sum the volume
Dim TotalVol As Variant

'I'm using reference equations so I have to select a cell to reference from
Range("A1").Select

'The while statement is looking at the 1st column.   At the end of each look I am moving down the selected cell by 1 until it reaches the end where the empty cell=0 and becomes false
While ActiveCell.Value > 0
    
    'I could have made this more versitile and set a var equal to the first value but all the sheets use <ticker>
    If ActiveCell.Value = "<ticker>" Then
        ActiveCell.Offset(0, 9).Value = "ticker"
        ActiveCell.Offset(0, 10).Value = "Yearly Change"
        ActiveCell.Offset(0, 11).Value = "Percent Change"
        ActiveCell.Offset(0, 12).Value = "Total Stock Volume"
        ActiveCell.Offset(0, 16).Value = "Ticker"
        ActiveCell.Offset(0, 17).Value = "Value"
        ActiveCell.Offset(1, 15).Value = "Greatest % Increase"
        ActiveCell.Offset(2, 15).Value = "Greatest % Decrease"
        ActiveCell.Offset(3, 15).Value = "Greatest Total Volume"
        
        'I had to put the following code in this If statement because the ActiveCell.Offset(-1,5).Value (in the next section) returns a string ends in an error
        Cells(2, 10).Value = ActiveCell.Offset(1, 0).Value
        Count = Count + 1
        
        'I'm setting CellNameTemp = the first ticker value
        CellNameTemp = ActiveCell.Offset(1, 0).Value
        
        
        FirstOpen = ActiveCell.Offset(1, 2).Value
        
        'Moves the selection down one row
        ActiveCell.Offset(1, 0).Select
        
        TotalVol = TotalVol + ActiveCell.Offset(1, 6).Value
        
        
    'This statementbecomes true when the ticker switches.  This assumes the data is ordered by first Ticker and then Date!!
    ElseIf ActiveCell.Value <> CellNameTemp Then
        'This should calculate previous close and open
        LastClose = ActiveCell.Offset(-1, 5).Value
        YearlyChange = LastClose - FirstOpen
        Cells(Count - 1, 11).Value = YearlyChange
        'There are a couple Tickers that don't have any value in the beginning which causes a UND.
        If FirstOpen <> 0 Then
            Cells(Count - 1, 12).Value = YearlyChange / FirstOpen
        Else
            Cells(Count - 1, 12).Value = 0
        End If
        Cells(Count - 1, 13).Value = TotalVol
        
        
        
        Cells(Count, 10).Value = ActiveCell.Value
        Count = Count + 1
        CellNameTemp = ActiveCell.Value
        
       
             
        
        'redefining open for this set
        FirstOpen = ActiveCell.Offset(0, 2).Value
        TotalVol = ActiveCell.Offset(0, 6).Value
    Else
        TotalVol = TotalVol + ActiveCell.Offset(0, 6).Value
    End If
        
        
    
'This is important for the while.  It moves the selection down one row so we eventually come to 0 and exit the while loop
ActiveCell.Offset(1, 0).Select
Wend

'Finishes the last Row
LastClose = ActiveCell.Offset(-1, 5).Value
YearlyChange = LastClose - FirstOpen

'Count-1 because I made count 2 to start to smoothe out an earlier version and then changed directions and had to recorrect it.  As Code was working I left it.
Cells(Count - 1, 11).Value = YearlyChange
Cells(Count - 1, 12).Value = YearlyChange / FirstOpen
Cells(Count - 1, 13).Value = TotalVol





'Setting variables to keep tabs on the ticker name and the current max and min values
Dim GreatestIncrease As Variant
Dim TickerIncrease As Variant
Dim GreatestDecrease As Variant
Dim TickerDecrease As Variant
Dim MaxVol As Variant
Dim TickerVol As Variant


Range("J2").Select

GreatestIncrease = ActiveCell.Offset(0, 2).Value
TickerIncrease = ActiveCell.Value
GreatestDecrease = ActiveCell.Offset(0, 2).Value
TickerDecrease = ActiveCell.Value
MaxVol = ActiveCell.Offset(0, 3).Value
TickerVol = ActiveCell.Value


Range("J3").Select
'This loop works with the unique list of ticker values to find the max and min percent
While ActiveCell.Value > 0
    
    
    If ActiveCell.Offset(0, 2).Value > GreatestIncrease Then
        GreatestIncrease = ActiveCell.Offset(0, 2).Value
        TickerIncrease = ActiveCell.Value
               
    ElseIf ActiveCell.Offset(0, 2).Value < GreatestDecrease Then
        GreatestDecrease = ActiveCell.Offset(0, 2).Value
        TickerDecrease = ActiveCell.Value
        
    Else
        
    End If
    
    If ActiveCell.Offset(0, 3).Value > MaxVol Then
        MaxVol = ActiveCell.Offset(0, 3).Value
        TickerVol = ActiveCell.Value
    Else
    End If
        
        
    

ActiveCell.Offset(1, 0).Select
Wend





Range("Q2").Value = TickerIncrease
Range("R2").Value = GreatestIncrease
Range("Q3").Value = TickerDecrease
Range("R3").Value = GreatestDecrease
Range("Q4").Value = TickerVol
Range("R4").Value = MaxVol





Next AllSheets

End Sub

