Sub stock_analysis()
Dim ws As Worksheet
   


'Define all variables
    
    Dim ticker As String
    Dim ttlStockVol As Double
    Dim qtrChange As Double
    Dim startData As Long
    Dim lastRow As Long
    Dim percentChange As Double
    Dim i As Long
    
  For Each ws In ThisWorkbook.Worksheets
    
    'Set header rows
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Set Initial Values
 
    ttlStockVol = 0
    qtrChange = 0
    percentChange = 0
    startData = 2
    
    'Row number of last row of data
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Summarize each each ticker symbol qrtly change %change and total volume if ticker changes print results
    
        For i = 2 To lastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            openPrice = ws.Cells(i, 3).Value
       	    ttlStockVol = ws.Cells(i, 7).Value

                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
        
                    closePrice = ws.Cells(i, 6).Value
                    qtrChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = ((closePrice - openPrice) / openPrice)
                    Else
                        percentChange = "0.00%"
                        
                    End If
                    
		    ttlStockVol = ttlStockVol + ws.Cells(i, 7).Value
    
                ws.Range("I" & startData).Value = ticker
                ws.Range("J" & startData).Value = qtrChange
                If (qtrChange > 0) Then
                    ws.Range("J" & startData).Interior.ColorIndex = 4
                    
                ElseIf (qtrChange <= 0) Then
                    ws.Range("J" & startData).Interior.ColorIndex = 3
                End If
                
                ws.Range("K" & startData).Value = percentChange
                ws.Range("K:K").NumberFormat = "0.00%"
                ws.Range("L" & startData).Value = ttlStockVol
               
    
    'Reset values
                    startData = startData + 1
                    qtrChange = 0
                    ttlStockVol = 0
            
                Else
                    ttlStockVol = ttlStockVol + ws.Cells(i, 7).Value
        
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
maxIncrease = WorksheetFunction.Max(ws.Range("K:K"))
maxDecrease = WorksheetFunction.Min(ws.Range("K:K"))
maxVolume = WorksheetFunction.Max(ws.Range("L:L"))

ws.Range("Q2").Value = maxIncrease
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = maxDecrease
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = maxVolume

'Match maximums to correct ticker and print
increaseTicker = Application.WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0)
decreaseTicker = Application.WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0)
maxVolumeTicker = Application.WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0)

ws.Range("P2").Value = ws.Range("I" & increaseTicker)
ws.Range("P3").Value = ws.Range("I" & decreaseTicker)
ws.Range("P4").Value = ws.Range("I" & maxVolumeTicker)
   
    Next ws
End Sub