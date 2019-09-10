Attribute VB_Name = "Module1"
Sub StockAnalysis()

'Loop for each worksheet
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'Editing cells with descriptive text
Range("J1").Value = "<Ticker>"
Range("K1").Value = "<Open>"
Range("L1").Value = "<Close>"
Range("M1").Value = "<% Change>"
Range("N1").Value = "<Total Volume>"
Range("P2").Value = "<Greatest % Increase>"
Range("P3").Value = "<Greatest % Decrease>"
Range("P4").Value = "<Greatest Total Volume>"
Range("Q1").Value = "<Ticker>"
Range("R1").Value = "<Value>"

'Initial variable declaration
Dim TickerTotal, TotalVol As LongLong
Dim VolumeTotalRow, col, BigInc, BigDec As Integer
Dim RowNum, counter As Long
Dim ITick, DTick, VolTick As String

RowNum = Cells(Rows.Count, 1).End(xlUp).Row
counter = 2
VolumeTotalRow = 2
TotalVol = 0
BigInc = 0
BigDec = 0
ws_num = ThisWorkbook.Worksheets.Count

'Finds the yearly open and close for each ticker
For i = 2 To RowNum
    If Cells(i, 1).Value <> Cells(i - 1, 1) And Cells(i, 1).Value = Cells(i + 1, 1) Then
        Cells(counter, 11).Value = Cells(i, 3).Value
    ElseIf Cells(i, 1).Value = Cells(i - 1, 1) And Cells(i, 1).Value <> Cells(i + 1, 1) Then
        Cells(counter, 12).Value = Cells(i, 6).Value
        Cells(counter, 10).Value = Cells(i, 1).Value
        counter = counter + 1
    End If
Next i

col = Cells(Rows.Count, 10).End(xlUp).Row

'Calculates percent change
For j = 2 To col
    If Cells(j, 12).Value <> 0 And Cells(j, 11).Value <> 0 Then
        Cells(j, 13).Value = (((Cells(j, 12).Value - Cells(j, 11).Value) / Cells(j, 11).Value) * 100)
    Else
        Cells(j, 13).Value = 0
    End If
Next j

'Calculates total volume for each ticker
For k = 2 To RowNum
    If Cells(k + 1, 1).Value <> Cells(k, 1).Value Then
        TickerTotal = TickerTotal + Cells(k, 7).Value
        Cells(VolumeTotalRow, 14).Value = TickerTotal
        VolumeTotalRow = VolumeTotalRow + 1
        TickerTotal = 0
    Else
        TickerTotal = TickerTotal + Cells(k, 7).Value
    End If
Next k

'Conditional green cell for positive %change, red cell for negative %change
For l = 2 To col
    If Cells(l, 13).Value > 0 Then
        Cells(l, 13).Interior.ColorIndex = 4
    ElseIf Cells(l, 13).Value < 0 Then
        Cells(l, 13).Interior.ColorIndex = 3
    End If
Next l

'Finds the biggest %increase and %decrease for all tickers in each year
For m = 2 To col
    If Cells(m, 13).Value >= BigInc Then
        BigInc = Cells(m, 13).Value
        ITick = Cells(m, 10).Value
    ElseIf Cells(m, 13).Value <= BigDec Then
        BigDec = Cells(m, 13).Value
        DTick = Cells(m, 10).Value
    End If
Next m

'Finds the biggest total volume for all tickers in each year
For n = 2 To col
    If Cells(n, 14).Value >= TotalVol Then
        TotalVol = Cells(n, 14).Value
        VolTick = Cells(n, 10).Value
    End If
Next n

Range("R2").Value = BigInc
Range("R3").Value = BigDec
Range("R4").Value = TotalVol
Range("Q2").Value = ITick
Range("Q3").Value = DTick
Range("Q4").Value = VolTick

Next ws

End Sub
