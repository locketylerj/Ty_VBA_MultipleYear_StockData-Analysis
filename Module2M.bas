Attribute VB_Name = "Module2"
Sub HMultipleYearStockAnalysis()

For Each ws In Worksheets

Dim TikTotal As Double
Dim i As Long
Dim j As Long
Dim YearlyChange As Double
Dim PercentChange As Long
Dim lasttrow As Long
Dim start As Long


j = 0
TikTotal = 0
start = 2
YearlyChange = 0
PercentChange = 0

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "TotalStockVolume"
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "PercentChange"

    For i = 2 To lastrow
    
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            TikTotal = TikTotal + ws.Cells(i, 7).Value
            
            'Handle zero total volume (if there is zero total volume then
            'there is zero percent change and yearly change for that stock?
            If TikTotal = 0 Then
                ws.Range("H" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("I" & 2 + j).Value = 0
                ws.Range("J" & 2 + j).Value = "%" & 0
                ws.Range("K" & 2 + j).Value = 0
            Else
            
            'Find first non zero starting value
            If ws.Cells(start, 3) = 0 Then
                For find_value = start To i
                    If ws.Cells(find_value, 3).Value <> 0 Then
                        start = find_value
                        Exit For
                    End If
                Next find_value
            End If
                
                
            YearlyChange = (Cells(i, 6) - ws.Cells(start, 3))
            PercentChange = Round((YearlyChange / ws.Cells(start, 3) * 100), 2)
            
            'start of the next stock ticker
            
            start = i + 1
        
            ws.Range("H" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("I" & 2 + j).Value = TikTotal
            ws.Range("J" & 2 + j).Value = YearlyChange
            ws.Range("K" & 2 + j).Value = "%" & PercentChange
            
            'Color cells green if yearly change is >0 or red if <0 using Case structure
        
            Select Case YearlyChange
                Case Is > 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            End Select
         End If
            'Reset the values for the next, different stock ticker
            
            TikTotal = 0
            
            'Move to the next row for the next, different stock ticker
            j = j + 1
            
            YearlyChange = 0
        
        Else
        
            'If ticker is the same row to row then just add the volumes
            TikTotal = TikTotal + ws.Cells(i, 7).Value
         
        End If
            
    Next i
  
Next ws


End Sub



