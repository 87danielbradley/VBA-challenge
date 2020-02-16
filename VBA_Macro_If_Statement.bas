Attribute VB_Name = "Module2"
Sub Master()
Dim AllSheets As Worksheet
For Each AllSheets In ThisWorkbook.Worksheets
AllSheets.Select

Dim DSheet As Worksheet
Dim LastCol As Long
Dim LastRow As Long
Dim LittleRow As Long

''''
''''Double check active.worksheet
Set DSheet = ActiveSheet
LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column


'FirstOpen holds the Opening ticker value
Dim FirstOpen As Double
'LastClose holds the value for the Final Stock Close
Dim LastClose As Double
'Yearly Change = LastClose - First Open
Dim YearlyChange As Double


Dim Count As Integer
Count = 2


Dim CellNameTemp As String
 
Dim TotalVol As Variant
TotalVol = 0
 
Range("A1").Select
ActiveCell.Offset(0, 9).Value = "ticker"
ActiveCell.Offset(0, 10).Value = "Yearly Change"
ActiveCell.Offset(0, 11).Value = "Percent Change"
ActiveCell.Offset(0, 12).Value = "Total Stock Volume"
ActiveCell.Offset(0, 16).Value = "Ticker"
ActiveCell.Offset(0, 17).Value = "Value"
ActiveCell.Offset(1, 15).Value = "Greatest % Increase"
ActiveCell.Offset(2, 15).Value = "Greatest % Decrease"
ActiveCell.Offset(3, 15).Value = "Greatest Total Volume"


CellNameTemp = ActiveCell.Offset(1, 0).Value

ActiveCell.Offset(1, 9).Value = CallNametemp

FirstOpen = ActiveCell.Offset(1, 2).Value


For i = 2 To LastRow

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Cells(Count, 10).Value = CellNameTemp
        CellNameTemp = Cells(i + 1, 1).Value
        TotalVol = TotalVol + Cells(i, 7).Value
        LastClose = Cells(i, 6).Value
        YearlyChange = LastClose - FirstOpen
        Cells(Count, 11).Value = YearlyChange
        
        'Conditional Formating
        If YearlyChange >= 0 Then
            
            Cells(Count, 11).Interior.ColorIndex = 4
        Else
            
            Cells(Count, 11).Interior.ColorIndex = 3
        End If
        
        
        
        Cells(Count, 13).Value = TotalVol
        
        If FirstOpen <> 0 Then
            Cells(Count, 12).Value = YearlyChange / FirstOpen
        Else
            Cells(Count, 12).Value = 0
        End If
        
        FirstOpen = Cells(i + 1, 3).Value
        Count = Count + 1
        TotalVol = 0
    Else
        TotalVol = TotalVol + Cells(i, 7).Value
    End If
        
 
  
Next i

'Finishes the last Row
LastClose = Cells(i - 1, 5).Value
YearlyChange = LastClose - FirstOpen

'Count-1 because I made count 2 to start to smoothe out an earlier version and then changed directions and had to recorrect it.  As Code was working I left it.
If YearlyChange >= 0 Then
    Cells(Count - 1, 11).Value = YearlyChange
    Cells(Count - 1, 11).Interior.ColorIndex = 4
Else
    Cells(Count - 1, 11).Value = YearlyChange
    Cells(Count - 1, 11).Interior.ColorIndex = 3
End If

If FirstOpen <> 0 Then
    Cells(Count - 1, 12).Value = YearlyChange / FirstOpen
Else
    Cells(Count - 1, 12).Value = 0
End If
Cells(Count - 1, 13).Value = TotalVol





'Setting variables to keep tabs on the ticker name and the current max and min values
Dim GreatestIncrease As Double
Dim TickerIncrease As String
Dim GreatestDecrease As Double
Dim TickerDecrease As String
Dim MaxVol As Variant
Dim TickerVol As String


Range("J2").Select

GreatestIncrease = ActiveCell.Offset(0, 2).Value
TickerIncrease = ActiveCell.Value
GreatestDecrease = ActiveCell.Offset(0, 2).Value
TickerDecrease = ActiveCell.Value
MaxVol = ActiveCell.Offset(0, 3).Value
TickerVol = ActiveCell.Value


LittleRow = DSheet.Cells(Rows.Count, 10).End(xlUp).Row
For j = 3 To LittleRow
    
    
    If Cells(j, 12).Value > GreatestIncrease Then
        GreatestIncrease = Cells(j, 12).Value
        TickerIncrease = Cells(j, 10).Value
               
    ElseIf Cells(j, 12).Value < GreatestDecrease Then
        GreatestDecrease = Cells(j, 12).Value
        TickerDecrease = Cells(j, 10).Value
        
    Else
        
    End If
    
    If Cells(j, 13).Value > MaxVol Then
        MaxVol = Cells(j, 13).Value
        TickerVol = Cells(j, 10).Value
    Else
    End If
        

Next j





Range("Q2").Value = TickerIncrease
Range("R2").Value = GreatestIncrease
Range("Q3").Value = TickerDecrease
Range("R3").Value = GreatestDecrease
Range("Q4").Value = TickerVol
Range("R4").Value = MaxVol





Next AllSheets

End Sub

