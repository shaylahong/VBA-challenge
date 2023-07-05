VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub module_2():

'loop labels through each worksheet

For Each ws In Worksheets
    Dim i As Double
    Dim k As Double
    Dim y As Double
    Dim m As Double
    Dim LastRow As Long
    Dim total_stock_volume As Double
    Dim rangeToFormat As Range
    Dim rangeToFormat2 As Range
    Dim FormatCondition As FormatCondition

    'y is the row where opening_price sits
    y = 2

    'k is the column where the yearly change value sits
    k = 2

    'labels
    ws.Range("I1,Q1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("R1").Value = "Value"


    'ticker symbol
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    total_stock_volume = 0
    For i = 2 To LastRow
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(k, 9).Value = ws.Cells(i, 1).Value
        
            ' yearly change from opening_price and closing_price
            Dim opening_price As Double
            Dim closing_price As Double

            opening_price = ws.Cells(y, 3).Value
            closing_price = ws.Cells(i, 6).Value
            ws.Range("J" & k).Value = closing_price - opening_price
        
            ' percent change from opening_price and closing_price
            m = k
            ws.Range("K" & m).Value = ((closing_price - opening_price) / opening_price)

            'total stock volume
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            ws.Range("L" & m).Value = total_stock_volume
            y = i + 1
            k = k + 1
            total_stock_volume = 0
    
        Else
            total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
        End If

    Next i
 
        ' Conditional formatting for yearly change
        Set rangeToFormat = ws.Range("J2:J" & LastRow)
        Set FormatCondition = rangeToFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0")
        FormatCondition.Interior.Color = vbGreen

        Set FormatCondition = rangeToFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        FormatCondition.Interior.Color = vbRed

        ' conditional formatting for percent change
        Set rangeToFormat2 = ws.Range("K2:K" & LastRow)
        Set FormatCondition = rangeToFormat2.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0")
        FormatCondition.Interior.Color = vbGreen
    
        Set FormatCondition = rangeToFormat2.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        FormatCondition.Interior.Color = vbRed

    'bonus: finding greatest percent increase/decrease
    'increase
    Dim max_increase As Double
    Dim max_increase_row As Long
    Dim current_increase As Double
    Dim ticker_increase As String

    max_increase_row = 2
    max_increase = ws.Cells(max_increase_row, "K").Value
    ticker_increase = ws.Cells(max_increase_row, "I").Value
    
    'decrease
    Dim max_decrease As Double
    Dim max_decrease_row As Long
    Dim current_decrease As Double
    Dim ticker_decrease As String

    max_decrease_row = 2
    max_decrease = ws.Cells(max_decrease_row, "K").Value
    ticker_decrease = ws.Cells(max_decrease_row, "I").Value

    'total stock volume
    Dim max_total_volume As Double
    Dim max_total_volume_row As Long
    Dim current_total_volume As Double
    Dim ticker_total_volume As String

    max_total_volume_row = 2
    max_total_volume = ws.Cells(max_total_volume_row, "L").Value
    ticker_total_volume = ws.Cells(max_total_volume_row, "I").Value


        For i = 3 To LastRow
        'increase
            current_increase = ws.Cells(i, "K").Value
            If current_increase > max_increase Then
                max_increase = current_increase
                max_increase_row = i
                ticker_increase = ws.Cells(i, "I").Value
        
            End If
        
        'decrease
            current_decrease = ws.Cells(i, "K").Value
            If current_decrease < max_decrease Then
                max_decrease = current_increase
                max_decrease_row = i
                ticker_decrease = ws.Cells(i, "I").Value
        
            End If
        
        'total stock volume
            current_total_volume = ws.Cells(i, "L").Value
            If current_total_volume > max_total_volume Then
                max_total_volume = current_total_volume
                max_total_volume_row = i
                ticker_total_volume = ws.Cells(i, "I").Value
        
            End If
        
        Next i
    

    'increase
    ws.Range("R2").Value = max_increase
    ws.Range("R2").NumberFormat = "0.00%"
    ws.Range("Q2").Value = ticker_increase
    
    'decrease
    ws.Range("R3").Value = max_decrease
    ws.Range("R3").NumberFormat = "0.00%"
    ws.Range("Q3").Value = ticker_decrease
    
    'total stock volume
    ws.Range("R4").Value = max_total_volume
    ws.Range("Q4").Value = ticker_total_volume

        Next ws

End Sub


