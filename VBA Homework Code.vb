Sub stockissues()

For Each ws In Worksheets


'data
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("o1").Value = "Stock Statistics"
ws.Range("o2").Value = "Greatest Percent Increase"
ws.Range("o3").Value = "Greatest Percent Decrease"
ws.Range("o4").Value = "Greatest Total Volume"


'output
'col 9 - ticker
'col 10 - yearly change
'col 11 - percent change
'col 12 - total stock volume
Dim dataRow As Long
Dim outputRow As Long
Dim openPrice As Double
Dim totalStockVolume As Double
Dim closePrice As Double
Dim PriceChange As Double
Dim PercentIncrease As Double
Dim PercentDecrease As Double
Dim Totalvolume As Double
Dim tickerincrease As String
Dim tickerdecrease As String
Dim tickervolume As String



openPrice = ws.Range("C2").Value
outputRow = 2
totalStockVolume = 0

For dataRow = 2 To (ws.Range("A2").End(xlDown).Row)
If ws.Cells(dataRow, 1).Value <> ws.Cells(dataRow + 1, 1).Value Then
'now at the edge
'add what is in Col to G to the total

'grab the closing price in column F
closePrice = ws.Cells(dataRow, 6).Value
If openPrice = 0 Then
            ws.Cells(outputRow, 11).Value = "Not Available"
           
Else
PriceChange = (closePrice - openPrice)

ws.Cells(outputRow, 11).Value = PriceChange / openPrice
totalStockVolume = totalStockVolume + ws.Cells(dataRow, 7).Value
End If

ws.Cells(outputRow, 10).Value = closePrice - openPrice

ws.Cells(outputRow, 12).Value = totalStockVolume

'Ticker
ws.Cells(outputRow, 9).Value = ws.Cells(dataRow, 1).Value

If ws.Cells(outputRow, 10).Value > 0 Then
         
ws.Cells(outputRow, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(outputRow, 10).Value < 0 Then
            
ws.Cells(outputRow, 10).Interior.ColorIndex = 3
            Else
               
ws.Cells(outputRow, 10).Interior.ColorIndex = 0
            End If
            
ws.Cells(outputRow, 11).NumberFormat = "0.00%"




totalStockVolume = 0


outputRow = outputRow + 1

'then update the new open price to be the open price of the next row
openPrice = ws.Cells(dataRow + 1, 3).Value

Else

totalStockVolume = totalStockVolume + ws.Cells(dataRow, 7).Value



End If

Next dataRow


'BONUS Part

    Sheets(Array("2016", "2015", "2014")).Select
    Sheets("2016").Activate
    Columns("O:O").EntireColumn.AutoFit
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-6])"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "=MIN(C[-6])"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=MAX(C[-5])"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[1],C[-5],C[-7],0,0)"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P3"), Type:=xlFillDefault
    Range("P2:P3").Select
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=XLOOKUP(RC[1],C[-4],C[-7],0,0)"
    Range("O1:Q4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select





Next ws



End Sub

