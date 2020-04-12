Attribute VB_Name = "Module1"
Sub Stock_Analysis()

'Iterate through each worksheet in the document
For Each ws In Worksheets

'Row value for loop
Dim i As Long

'Row value for results chart
Dim j As Integer

'Other variables
Dim Ticker As String
Dim Opening_Value As Double
Dim Closing_Value As Double
Dim Yearly_Change As Double
Dim Stock_Volume As Double

'Results chart headers
ws.Range("I" & 1).Value = "Ticker"
ws.Range("J" & 1).Value = "Yearly Change"
ws.Range("K" & 1).Value = "Percent Change"
ws.Range("L" & 1).Value = "Total Stock Volume"

'Initialize chart row and first stock opening value
j = 2
Opening_Value = ws.Range("C2").Value

'For loop to iterate until the last populated row is reached
For i = 2 To (ws.Range("A1").End(xlDown).Row)

'Designates the last row for a particular ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
'Provides ticker ID, closing value, and performs final stock volume calculation
        Ticker = ws.Cells(i, 1).Value
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        Closing_Value = ws.Cells(i, 6).Value
        
'Populates results chart with ticker and yearly change data (with color formatting for yearly gain/loss
        ws.Range("I" & j).Value = Ticker
        ws.Range("J" & j).Value = Closing_Value - Opening_Value
            If ws.Range("J" & j).Value >= 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 4
            Else
                ws.Range("J" & j).Interior.ColorIndex = 3
            End If

'Populates the percent change result for the ticker (entering N/A if the opening value was 0 to avoid a zero in the denominator)
        If Opening_Value = 0 Then
            ws.Range("K" & j).Value = "N/A"
        Else
            ws.Range("K" & j).Value = FormatPercent((ws.Range("J" & j).Value / Opening_Value), 2)
        End If
        
'Populates the total stock volume result for the ticker
        ws.Range("L" & j).Value = Stock_Volume
        
'Starts new results row, resets stock volume variable, and defines the new opening value
        j = j + 1
        Stock_Volume = 0
        Opening_Value = ws.Cells(i + 1, 3).Value

'If the ticker in the next row is the same, then just add the stock value and move on
    Else
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
    End If
    
Next i

'Move to next worksheet
Next ws

End Sub
