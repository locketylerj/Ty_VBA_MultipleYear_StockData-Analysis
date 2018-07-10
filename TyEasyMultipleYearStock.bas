Attribute VB_Name = "Module1"
Sub TyEasyMultipleYearStockData()

'Used a combination of coding from Wellsfargo pt1 VBA 3 in class activity _
 and solved credit card challenge in-class activity.
'Dim TickerSumTable As Integer
'Dim lastrow As Long
'Dim ws As Worksheet
'TickerSumTable = 2


Dim TickerTotal As Double
TickerTotal = 0

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Create Title Rows
Range("I1").Value = "Ticker"
Range("J1").Value = "TotalStockVolume"

j = 0

    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            TickerTotal = TickerTotal + Cells(i, 7).Value
        
        'print the ticker symbol
        
            Range("I" & 2 + j).Value = Cells(i, 1).Value
        
        'print the total
        
            Range("J" & 2 + j).Value = TickerTotal
        
        'Reset the total value and move to the next row.
            TickerTotal = 0
        
            j = j + 1
        
        
        Else
        
            TickerTotal = TickerTotal + Cells(i, 7).Value
        
        
        End If
        
    
    Next i
    
   
End Sub
