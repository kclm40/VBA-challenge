Attribute VB_Name = "Module1"


Sub VBA_Challenge()

'Set all variables

'set a name for the dimension ticker

    Dim Ticker As String

'set opening price point
    Dim opening_price As Double

'set closing price point
    Dim closing_price As Double

'set yearly change
    Dim Yearly_Change As Double
    
 'set stock volume
    Dim stock_volume As Double
    
'set % chg
    Dim Percent_Change As Double


'set name for summary

    Dim Summary As Integer
    
    
    Dim ws As Worksheet
    
'------------------------------------------------------------
    
'Loop

For Each ws In Worksheets

'set summary headings

    ws.Range("$I$1").Value = "Ticker"
    
    ws.Range("$J$1").Value = "Yearly Change"
    
    ws.Range("$K$1").Value = "Percent Change"
    
    ws.Range("$L$1").Value = "Total Stock Volume"
    
'set intigers
    Summary = 2
    last = 1
    stock_volume = 0

'set last row
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    For i = 2 To lastRow
    

'check to see if we are in the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker = ws.Cells(i, 1).Value
    
        last = last + 1
    
'get values of open & close price

    opening_price = ws.Cells(last, 3).Value
    closing_price = ws.Cells(i, 6).Value
    
'sum stock volume
For j = last To i

    stock_volume = stock_volume + ws.Cells(j, 7).Value
    
Next j

'calculate yearly change & percent change
    If opening_price = 0 Then
        Percent_Change = closing_price
    
    Else
    
    Yearly_Change = closing_price - opening_price
    Percent_Change = Yearly_Change / opening_price
    
  End If
    

'add the ticker name to the summary

    ws.Range("I" & Summary).Value = Ticker

'add stock volume to the summary

    ws.Range("L" & Summary).Value = stock_volume
    
'add yearlty change to the summary
    ws.Range("J" & Summary).Value = Yearly_Change

'add percent_change to the summary
    ws.Range("K" & Summary).Value = Percent_Change
    
 'format percent change
    ws.Range("K" & Summary).NumberFormat = "0.00%"

'--------------------------------
    
    
    Summary = Summary + 1
    
    
    Yearly_Change = 0
    Percent_Change = 0
    stock_volume = 0
    
    last = i
    

End If

Next i


'---------------------------------------------------------------

'conditional formatting

condlastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row

For k = 2 To condlastrow

    If ws.Cells(k, 10) > 0 Then
        ws.Cells(k, 10).Interior.ColorIndex = 4
    
    Else
    
        ws.Cells(k, 10).Interior.ColorIndex = 3
    
    End If
    
    Next k


Next ws

End Sub



