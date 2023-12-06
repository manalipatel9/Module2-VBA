VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub newloop()

    Dim tickername As String
    Dim tickervolume As Double
    Dim sum_row As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    sum_row = 2
    open_price = Cells(2, 3).Value

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
   For i = 2 To Lastrow
   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    tickername = Cells(i, 1).Value
    tickername = tickervolume + Cells(i, 7).Value
    
    Range("I" & sum_row).Value = tickername
    Range("L" & sum_row).Value = tickervolume
    
    close_price = Cells(i, 6).Value
    
    yearly_change = (close_price - open_price)
    
    Range("J" & sum_row).Value = yearly_change
    
    If (open_price = 0) Then
    
        percent_change = 0
        
    Else
    
        percent_change = yearly_change / open_price
    
    End If
    
    Range("K" & sum_row).Value = percent_change
    Range("K" & sum_row).NumberFormat = "0.00%"
    
    sum_row = sum_row + 1
    
    tickervolume = 0
    
    open_price = Cells(i + 1, 3)
    
    Else
    
    tickervolume = tickervolume + Cells(i, 7).Value
    
    End If
    
Next i

lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow_summary_table
    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 10
Else
    Cells(i, 10).Interior.ColorIndex = 3
End If

Next i

Range("P1").Value = "Ticker"
Range("O1").Value = "Value"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"


GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

For a = 2 To Lastrow

If Cells(a, 11).Value > GreatestIncrease Then
    GreatestIncrease = Cells(a, 11).Value
    Range("Q2").Value = GreatestIncrease
    Range("Q2").Style = "Percent"
    Range("Q2").NumberFormat = "0.00%"
    Range("P2").Value = Cells(a, 9).Value
    
End If

Next a

For b = 2 To Lastrow

    If Cells(b, 11).Value < GreatestDecrease Then
    GreatestDecrease = Cells(b, 11).Value
    Range("Q3").Value = GreatestDecrease
    Range("Q3").Style = "Percent"
    Range("Q3").NumberFormat = "0.00"
    Range("P3").Value = Cells(b, 9).Value

End If
 Next b
 
For c = 2 To Lastrow

If Cells(c, 12).Value > GreatestVolume Then
GreatestVolume = Cells(c, 12).Value
Range("Q4").Value = GreatestVolume
Range("P4").Value = Cells(c, 9).Value

End If

Next c

Columns("A:Q").AutoFit

    
    
End Sub
