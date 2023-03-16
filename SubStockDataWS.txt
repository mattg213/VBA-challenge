Sub StockDataWS()

Dim ws As Worksheet

For Each ws In Worksheets
    
        Dim ticker As String
    
        Dim volume As Double
        volume = 0
        
        Dim year_open, year_close As Double
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Dim lastrow As Long
        
            lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
    
        For I = 2 To lastrow
            If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
                year_open = Cells(I, 3).Value
                Else
            End If
        
            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                ticker = Cells(I, 1).Value
                year_close = Cells(I, 6).Value
                volume = volume + Cells(I, 7).Value

                
                Cells(1, 10).Value = "Ticker"
                Cells(1, 11).Value = "Yearly Change"
                Cells(1, 12).Value = "Percent Change"
                Cells(1, 13).Value = "Total Volume"
                
                Range("J" & Summary_Table_Row).Value = ticker
                Range("K" & Summary_Table_Row).Value = (year_close - year_open)
                Range("L" & Summary_Table_Row).Value = (year_close - year_open) / year_close
                Range("M" & Summary_Table_Row).Value = volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                volume = 0
                
                Else
            
                volume = volume + Cells(I, 7).Value
                
            End If
            
        Next I
Next ws
End Sub

