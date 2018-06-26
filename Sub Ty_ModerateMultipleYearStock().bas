Attribute VB_Name = "Module1"
Sub Ty_ModerateMultipleYearStock()

Dim TikTotal As Double
Dim i As Long
Dim j As Long
Dim YearlyChange As Double
Dim PercentChange As Double
Dim lasttrow As Long

j = 0
TikTotal = 0


For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row


ws.Range("H1").Value = "Ticker"
ws.Range("I1").Value = "TotalStockVolume"
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "PercentChange"

    For i = 2 To lastrow
    
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            TikTotal = TikTotal + ws.Cells(i, 7).Value
            YearlyChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            PercentChange = Round((YearlyChange / ws.Cells(i, 3).Value * 100), 2)
        
        
            ws.Range("H" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("I" & 2 + j).Value = TikTotal
            ws.Range("J" & 2 + j).Value = YearlyChange
            ws.Range("K" & 2 + j).Value = PercentChange
            
        'Reset the values for the next, different stock ticker
        
        TikTotal = 0
        
        'Move to the next row for the next, different stock ticker
        j = j + 1
        
        Else
        
        'If ticker is the same row to row then just add the volumes
        TikTotal = TikTotal + ws.Cells(i, 7).Value
         
        End If
      'Color cells green if yearly change is >0 or red if <0 using Case structure
        
            Select Case YearlyChange
            Case Is > 0
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            End Select
            
    Next i
Next ws


End Sub


