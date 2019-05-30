Attribute VB_Name = "Module1"
Sub moderate()

Dim open_p As Double
Dim close_p As Double
Dim price_change As Double
Dim percent_change As Double
Dim summary_row As Integer

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
summary_row = 2
tot_vol = 0
k = 2
 
For i = 2 To Lastrow
  
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ticket_symbol = Cells(i, 1).Value
    open_p = Cells(k, 3).Value
    close_p = Cells(i, 6).Value
    price_change = close_p - open_p
    percent_change = (close_p - open_p) / open_p
    tot_vol = tot_vol + Cells(i, 7).Value
  
    Range("I" & summary_row).Value = ticket_symbol
    Range("J" & summary_row).Value = price_change
    Range("K" & summary_row).Value = percent_change
    Range("L" & summary_row).Value = tot_vol
    
    If price_change > 0 Then
    Range("J" & summary_row).Interior.ColorIndex = 4

    ElseIf price_change < 0 Then
    Range("J" & summary_row).Interior.ColorIndex = 3
 
    End If
    k = i + 1
    summary_row = summary_row + 1
    tot_vol = 0
   
   Else
  
    tot_vol = tot_vol + Cells(i, 7).Value

  End If
Next i

    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"

End Sub


