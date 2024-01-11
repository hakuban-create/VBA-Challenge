VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub getSumarry()

'-------------------------------------
'Looping through all sheets
'-------------------------------------
For Each ws In Worksheets

'-------------------------------------
'Adding column headers
'-------------------------------------
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
'-------------------------------------
'Declaring all variables
'-------------------------------------
Dim OpeningPrice As Variant: OpeningPrice = ws.Cells(2, 3)
Dim ClosingPrice As Variant:  ClosingPrice = 0
Dim TotalVolume As Variant: TotalVolume = 0
Dim GreatestIncrease As Variant: GreatestIncrease = 0
Dim GreatestIncreaseTicker As String
Dim GreatestDecrease As Variant:  GreatestDecrease = 0
Dim GreatestDecreaseTicker As String
Dim GreatestTotal As Variant: GreatestTotal = 0
Dim GreatestTotalTicker As String
Dim PercentChange As Variant
Dim CurrentTicker As String
Dim YearlyChange As Variant

'-------------------------------------
'Looping through each row
'-------------------------------------
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For Row = 2 To LastRow
'-------------------------------------
'Actions when ticker changes
'-------------------------------------
If (ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1)) Then
ClosingPrice = ws.Cells(Row, 6)
TotalVolume = TotalVolume + ws.Cells(Row, 7)
'-------------------------------------
'Writing final values
'-------------------------------------
TickerLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
CurrentTicker = ws.Cells(Row, 1)
ws.Cells(TickerLastRow + 1, 9).Value = CurrentTicker
YearlyChange = ClosingPrice - OpeningPrice
ws.Cells(TickerLastRow + 1, 10).Value = YearlyChange
PercentChange = (ClosingPrice - OpeningPrice) / OpeningPrice
ws.Cells(TickerLastRow + 1, 11).Value = PercentChange
ws.Cells(TickerLastRow + 1, 12).Value = TotalVolume
If (YearlyChange > 0) Then
ws.Cells(TickerLastRow + 1, 10).Interior.ColorIndex = 4
Else
ws.Cells(TickerLastRow + 1, 10).Interior.ColorIndex = 3
End If
'-------------------------------------
'Tracking greatest increase, decrease and total volume
'-------------------------------------
If (PercentChange > GreatestIncrease) Then
GreatestIncrease = PercentChange
GreatestIncreaseTicker = CurrentTicker
End If

If (PercentChange < GreatestDecrease) Then
GreatestDecrease = PercentChange
GreatestDecreaseTicker = CurrentTicker
End If

If (TotalVolume > GreatestVolume) Then
GreatestVolume = TotalVolume
GreatestVolumeTicker = CurrentTicker
End If

OpeningPrice = ws.Cells(Row + 1, 3)
TotalVolume = 0
Else
TotalVolume = TotalVolume + ws.Cells(Row, 7)
End If


Next Row
'-------------------------------------
'Updating the column formats
'-------------------------------------
ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
'-------------------------------------
'Writing Greatest Decrease, Increase and Volume
'-------------------------------------
ws.Range("P2").Value = GreatestIncreaseTicker
ws.Range("Q2").Value = GreatestIncrease
ws.Range("P3").Value = GreatestDecreaseTicker
ws.Range("Q3").Value = GreatestDecrease
ws.Range("P4").Value = GreatestVolumeTicker
ws.Range("Q4").Value = GreatestVolume

Next ws

End Sub


