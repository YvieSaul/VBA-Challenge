Sub stock_analysis()
   

'Define all variables
    Dim ticker As String
    Dim ttlStockVol As Double
    Dim qtrChange As Double
    Dim startData As Long
    Dim lastRow As Long
    Dim percentChange As Double
    Dim i As Long
    
    
    'Set header rows
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Set Initial Values
    
    ttlStockVol = 0
    qtrChange = 0
    percentChange = 0
    startData = 2
    
    'Row number of last row of data
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Summarize each each ticker symbol qrtly change %change and total volume if ticker changes print results
    
     For i = 2 To lastRow
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        openPrice = Cells(i, 3).Value
	ttlStockVol = Cells(i, 7).Value
       
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
        
                closePrice = Cells(i, 6).Value
                qtrChange = closePrice - openPrice
                
                If openPrice <> 0 Then
			percentChange = ((closePrice - openPrice) / openPrice)
		Else
			percentChange = "0.00%"
		End If
                ttlStockVol = ttlStockVol + Cells(i, 7).Value
                Range("I" & startData).Value = ticker
                Range("J" & startData).Value = qtrChange
                If (qtrChange > 0) Then
                    Range("J" & startData).Interior.ColorIndex = 4
                    
                ElseIf (qtrChange <= 0) Then
                    Range("J" & startData).Interior.ColorIndex = 3
                End If
                
                Range("K" & startData).Value = percentChange
                Range("K:K").NumberFormat = "0.00%"
                Range("L" & startData).Value = ttlStockVol
               
    
    'Reset values
                qtrChange = 0
                ttlStockVol = 0
                startData = startData + 1
            Else
                ttlStockVol = ttlStockVol + Cells(i, 7).Value

        
            
        End If
         
        
        
   Next i

 'Define variables
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double
Dim increaseTicker As String
Dim decreaseTicker As String
Dim maxVolumeTicker As String

'Find maximums and print
maxIncrease = WorksheetFunction.Max(Range("K:K"))
maxDecrease = WorksheetFunction.Min(Range("K:K"))
maxVolume = WorksheetFunction.Max(Range("L:L"))

Range("Q2").Value = maxIncrease
Range("Q2").NumberFormat = "0.00%"
Range("Q3").Value = maxDecrease
Range("Q3").NumberFormat = "0.00%"
Range("Q4").Value = maxVolume

'Match maximums to correct ticker and print
increaseTicker = Application.WorksheetFunction.Match(Range("Q2").Value, Range("K:K"), 0)
decreaseTicker = Application.WorksheetFunction.Match(Range("Q3").Value, Range("K:K"), 0)
maxVolumeTicker = Application.WorksheetFunction.Match(Range("Q4").Value, Range("L:L"), 0)

Range("P2").Value = Range("I" & increaseTicker)
Range("P3").Value = Range("I" & decreaseTicker)
Range("P4").Value = Range("I" & maxVolumeTicker)
   
End Sub

