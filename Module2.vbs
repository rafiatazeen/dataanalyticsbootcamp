Attribute VB_Name = "Module1"
Sub stock()
'loop through all sheets
For Each ws In Worksheets

'set an initial variable for holding the ticker
Dim ticker As String

'find the last row
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'keep track of the ticker location
Dim location As Integer
location = 2

'make ticker header
ws.Cells(1, 9).Value = "Ticker"

'set variable for opening price
Dim tickeropen As Double

'set variable for closing price
Dim tickerclose As Double

'set variable for yearly change
Dim yearlychange As Double

'make header for yearly change
ws.Cells(1, 10).Value = "Yearly Change"

'set variable for percent change
Dim percentchange As Double

'make header for percent change
ws.Cells(1, 11).Value = "Percent Change"

'set variable for total stock volume
Dim stockvolume As Double
stockvolume = 0

'make header for stock volume
ws.Cells(1, 12).Value = "Total Stock Volume"

'loop through all the tickers
For i = 2 To lastrow

'value for opening price
If ws.Cells(i, 2).Value = ws.Cells(2, 2).Value Then
tickeropen = ws.Cells(i, 3).Value
End If

'check to see if same ticker or not
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
tickerclose = ws.Cells(i, 6).Value
yearlychange = tickerclose - tickeropen
percentchange = yearlychange / tickeropen
stockvolume = stockvolume + ws.Cells(i, 7).Value

'figure out if yearly change is negative or positive and change cell colour accordingly
If yearlychange > 0 Then
ws.Range("J" & location).Interior.ColorIndex = 4
Else
ws.Range("J" & location).Interior.ColorIndex = 3
End If

'print ticker in column
ws.Range("I" & location).Value = ticker

'print yearly change in column
ws.Range("J" & location).Value = yearlychange

' percent change in column
ws.Range("K" & location).Value = percentchange

'print stock volume
ws.Range("L" & location).Value = stockvolume

'add one to the location
location = location + 1

'reset values
tickerclose = 0
yearlychange = 0
percentchange = 0
stockvolume = 0

Else
tickerclose = ws.Cells(i, 6).Value
yearlychange = tickerclose - tickeropen
percentchange = yearlychange / tickeropen
stockvolume = stockvolume + ws.Cells(i, 7).Value


End If

Next i

'format the percent change column
For i = 2 To lastrow
ws.Cells(i, 11).Style = "Percent"
Next i

'make table
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'set variable for greatest percentage increase
Dim greatest_increase As Double

For i = 2 To lastrow
    With ws.Cells(i, 11)
        If .Value > Max Then
        Max = .Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        End If
    End With
Next i

greatest_increase = Max

'format greatest percentage increase
ws.Cells(2, 17).Style = "Percent"

'print greatest percentage increase
ws.Cells(2, 17).Value = greatest_increase

Max = 0


'set variable for greatest percentage decrease
Dim greatest_decrease As Double

For i = 2 To lastrow
    With ws.Cells(i, 11)
        If .Value < Min Then
        Min = .Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        End If
    End With
Next i

greatest_decrease = Min

'format greatest percentage decrease
ws.Cells(3, 17).Style = "Percent"

'print greatest percentage decrease
ws.Cells(3, 17).Value = greatest_decrease

'find greatest total volume
Dim greatest_volume As Double
    For i = 2 To lastrow
        With ws.Cells(i, 12)
            If .Value > Max Then
                Max = .Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        End With
    Next i
greatest_volume = Max
ws.Cells(4, 17).Value = greatest_volume

Max = 0


ws.Columns("I:Q").AutoFit




Next ws

End Sub
