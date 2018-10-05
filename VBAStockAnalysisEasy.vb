Sub StockAnalysisEasy()
    Dim ws As Worksheet
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim VolumeRow As Integer

    For Each ws In Worksheets
    ws.Activate

    TotalVolume = 0

    VolumeRow = 2
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                Ticker = Cells(i, 1).Value
                TotalVolume = TotalVolume + Cells(i, 7).Value

                ws.Range("i" & VolumeRow).Value = Ticker

                ws.Range("j" & VolumeRow).Value = TotalVolume

                VolumeRow = VolumeRow + 1
                
                TotalVolume = 0

            Else
                TotalVolume = TotalVolume + Cells(i, 7).Value

            End If
              
        Next i

    Next ws

End Sub