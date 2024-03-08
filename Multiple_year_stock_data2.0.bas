Attribute VB_Name = "Module1"

Sub StockMarket()

'Loop Each Worksheet
Dim WS As Worksheet
For Each WS In Worksheets

'Define Variables
WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total Stock Volume"
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Total_Stock_Volume = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'Define Opening Price
Opening_Price = WS.Cells(2, 3).Value

For i = 2 To LastRow
  If WS.Cells(i + 1, 1).Value = WS.Cells(i, 1).Value Then
    Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
  ElseIf WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    ' Populate ticker symbol
    WS.Range("I" & Summary_Table_Row).Value = WS.Cells(i, 1).Value
    
    ' Calculate Yearly Change
    Closing_Price = WS.Cells(i, 6).Value
    WS.Range("J" & Summary_Table_Row).Value = Closing_Price - Opening_Price
    Yearly_Change = WS.Range("J" & Summary_Table_Row).Value
    WS.Range("K" & Summary_Table_Row).Value = FormatPercent(Yearly_Change / Opening_Price, 2)
    Opening_Price = WS.Cells(i + 1, 3).Value
    
     If Yearly_Change > 0 Then
       WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      Else
       WS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      End If

    ' Calculate Total Stock Volume
    Total_Stock_Volume = Total_Stock_Volume + WS.Cells(i, 7).Value
    WS.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    Summary_Table_Row = Summary_Table_Row + 1
    Total_Stock_Volume = 0
    End If
 Next i
 
 ' Bonus Questions
  WS.Range("O2").Value = "Greatest% Increase"
  WS.Range("O3").Value = "Greatest% Decrease"
  WS.Range("O4").Value = "Greatest Total Volumn"
  WS.Range("P1").Value = "Ticker"
  WS.Range("Q1").Value = "Value"
  
  Dim Row_Num As Integer
  Row_Num = Summary_Table_Row - 1
  Greatest_Increase = WS.Range("K2").Value
  Greatest_Decrease = WS.Range("K2").Value
  Greatest_Total = WS.Range("L2").Value
  
 ' calculate greatest total volumn
   For i = 3 To Row_Num
    If WS.Cells(i, 12).Value > Greatest_Total Then
    Greatest_Total = WS.Cells(i, 12).Value
    WS.Range("Q4").Value = Greatest_Total
    WS.Range("P4").Value = WS.Cells(i, 9).Value
    End If
    
    ' calculate greatest%
    If WS.Cells(i, 11).Value > Greatest_Increase Then
      Greatest_Increase = WS.Cells(i, 11).Value
      increase_ticker = i
    ElseIf WS.Cells(i, 11) < Greatest_Decrease Then
      Greatest_Decrease = WS.Cells(i, 11).Value
      decrease_ticker = i
    End If
   Next i
   WS.Range("Q2").Value = FormatPercent(Greatest_Increase, 2)
   WS.Range("P2").Value = WS.Cells(increase_ticker, 9).Value
   WS.Range("Q3").Value = FormatPercent(Greatest_Decrease, 2)
   WS.Range("P3").Value = WS.Cells(decrease_ticker, 9).Value
Next WS
End Sub
