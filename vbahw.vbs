Sub stocktotaler_easy()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim Summary_Table_Row As Integer



For Each ws In ThisWorkbook.Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"

    
   
    Summary_Table_Row = 2
       
        For i = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                  
                 ticker = ws.Cells(i, 1).Value
                   vol = ws.Cells(i, 7).Value
                
                
                    
                  ws.Cells(Summary_Table_Row, 9).Value = ticker
                    ws.Cells(Summary_Table_Row, 10).Value = vol
                 Summary_Table_Row = Summary_Table_Row + 1
        
                 vol = 0
        
             End If
        Next i
       
Next
End Sub