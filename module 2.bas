Attribute VB_Name = "Module1"
Sub multiple_year_stock_data()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        Dim i As Long
        Dim j As Long
        Dim tick As Long
        Dim percent As Double
        Dim grInc As Double
        Dim grDec As Double
        Dim grVol As Double
        Dim lastA As Long
        Dim lastI As Long
        
        initial1 = 0
        initial2 = 0
        intial3 = 0
        initial4 = 2
        
    
    Next
End Sub

