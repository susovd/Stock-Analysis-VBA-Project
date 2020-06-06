Attribute VB_Name = "Module1"
Sub samplestocks()
'keeping track of values as we iterate over rows

For Each ws In Worksheets
'choose worksheet
ws.Activate

'assign ticker variable for to loop over each ticker row
Dim ticker As String

'store individual ticker name
Dim tickername As String
'assign opening value to a ticker
Dim openValue As Double

'assign closing value to a ticker
Dim closeValue As Double

'assign yearly change variable
Dim Yearly_change As Double

'assign volume for a row
Dim currentVolume As Long

'add volumes for each stock
Dim cumulativeVolume As Double
cumulativeVolume = 0

'start from row 2
Dim i As Long
i = 2

'declare areas for summary
Dim summaryRow As Long
summaryRow = 2

'assign headers to results section
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"


'get first stock ticker
ticker = Cells(i, 1).Value

'set openValue to 0 for each ticker
'openValue = 0

'loop through all rows in active sheet until end of rows
Do While ticker <> ""
openValue = 0
If openValue = 0 Then
    openValue = Cells(i, 3).Value
End If

currentVolume = Cells(i, 7).Value
cumulativeVolume = cumulativeVolume + currentVolume
tickername = Cells(i, 1).Value
    
'identify last row of a ticker
If ticker <> Cells(i + 1, 1).Value Then
    Cells(summaryRow, 9).Value = ticker
    Cells(summaryRow, 12).Value = cumulativeVolume
    closeValue = Cells(i, 6).Value
    cumulativeVolume = 0
    
    'yearly change calculation
        Cells(summaryRow, 10) = closeValue - openValue
    
    'calculate percent change
        If openValue = 0 Then
            Cells(summaryRow, 11) = 0
        Else
            Cells(summaryRow, 11) = Cells(summaryRow, 10).Value / openValue
        End If
        
        
        'conditinal formatting for negative and positive outcome
                If Cells(summaryRow, 11).Value < 0 Then
                    Cells(summaryRow, 11).Interior.ColorIndex = 3
                Else
                    Cells(summaryRow, 11).Interior.ColorIndex = 4
                End If
                
        Do While Cells(summaryRow, 11).Value <> ""
        Cells(summaryRow, 11).Value = FormatPercent(Cells(summaryRow, 11).Value)
                
        
        'move to next row for next ticker summary
        summaryRow = summaryRow + 1
        Loop
End If
    i = i + 1
    
    'move to a new ticker
    ticker = Cells(i, 1).Value
Loop


'hard part
'declare j as a variable to loop through all rows in summary section
Dim j As Integer

'results start from 2nd row
j = 2

'find last row of summary
lastRowSummary = ws.Cells(Rows.Count, "I").End(xlUp).Row
 

'assign first ticker values for all items we need to present

'ticker with greatest percent increase
greatestPcntInc = Cells(2, 11).Value
greatestPctIncTicker = Cells(2, 9).Value

'ticker with greatest percent decrease
greatestPcntDec = Cells(2, 11).Value
greatestPctDecTicker = Cells(2, 9).Value

'ticker with greatest volume
greatestVol = Cells(2, 12).Value
greatestVolTicker = Cells(2, 9).Value

For j = 2 To lastRowSummary

'ticker with greatest percent increase
If Cells(j, 11).Value > greatestPcntInc Then
    greatestPcntInc = Cells(j, 11).Value
    greatestPcntIncTicker = Cells(j, 9).Value
End If

'ticker with greatest percent decrease
If Cells(j, 11).Value < greatestPcntDec Then
    greatestPcntDec = Cells(j, 11).Value
    greatestPcntDecTicker = Cells(j, 9).Value
End If

'ticker with greatest volume

If Cells(j, 12).Value > greatestVol Then
    greatestVol = Cells(j, 12).Value
    greatestVolTicker = Cells(j, 9).Value
End If

Next j





'Cells(summaryRow, 11).Value = FormatPercent(Cells(summaryRow, 11).Value)

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("P2").Value = greatesPcntIncTicker
Range("Q2").Value = FormatPercent(greatestPcntInc)
Range("P3").Value = greatestPcntDecTicker
Range("Q3").Value = FormatPercent(greatestPcntDec)
Range("P4").Value = greatestVolTicker
Range("Q4").Value = greatestVol

'Another way to find max value in a column
'Cells(2, 17).Value = Application.WorksheetFunction.Max(Range("k:k"))
'Cells(3, 17).Value = Application.WorksheetFunction.Min(Range("k:k"))
'Cells(4, 17).Value = Application.WorksheetFunction.Max(Range("l:l"))


'Cells(2, 16).Value = Application.WorksheetFunction.Max(Range("K:K")).Offset(, -1), doesn't work
'Cells(summaryRow, 11).Value = FormatPercent(Cells(summaryRow, 11).Value)




Next ws






End Sub



