VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarket():

Dim TickerSymbol As String

Dim TickerTotal As Double
TickerTotal = 0

Dim DailyStockStats As Integer
DailyStockStats = 2

Dim YearOpeningPrice As Double
YearOpeningPrice = Cells(2, 3).Value


    For i = 2 To 760192

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            TickerSymbol = Cells(i, 1).Value

            TickerTotal = TickerTotal + Cells(i, 7).Value

            Cells(DailyStockStats, 9).Value = TickerSymbol

            Cells(DailyStockStats, 12).Value = TickerTotal

            TickerTotal = 0

            YearClosingPrice = Cells(i, 6).Value

            YearlyChange = YearClosingPrice - YearOpeningPrice

            Range("J" & DailyStockStats).Value = YearlyChange

            Range("K" & DailyStockStats).Value = YearlyChange / YearOpeningPrice

            Range("K:K").NumberFormat = "0.00%"

            DailyStockStats = DailyStockStats + 1

            'If Cells(i, 10).Value < 0 Then

            'Cells(i, 10).Interior.ColorIndex = 3

            'If Cells(i, 10).Value >= 0 Then

            'Cells(i, 10).Interior.ColorIndex = 4

        Else

            TickerTotal = TickerTotal + Cells(i, 7).Value

        End If

    Next i

End Sub





