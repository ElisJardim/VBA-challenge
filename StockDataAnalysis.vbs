Attribute VB_Name = "StockDataAnalysis"


Sub StockDataAnalysis():

    'Define Variables
    Dim Ticker As String
    Dim OpenValue As Double
    Dim HighValue As Double
    Dim LowValue As Double
    Dim CloseValue As Double
    Dim Volume As Long
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
    Dim ResultIndex As Integer
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' Add the headers
        With ws
            .Cells(1, 9).Value = "Ticker"
            .Cells(1, 10).Value = "Yearly Change"
            .Cells(1, 11).Value = "Percent Change"
            .Cells(1, 12).Value = "Total Stock Volume"
        End With
     
        'Count rows
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        
        ResultIndex = 2
        IsFirstItem = True
        TotalVolume = 0
        For I = 2 To LastRow
            'Did Ticker change?
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                
                ' Set Ticker column in the results
                ws.Cells(ResultIndex, 9).Value = ws.Cells(I, 1).Value
                
                'Calculate, Save, and Format Yearly Change
                CloseValue = ws.Cells(I, 6).Value
                YearlyChange = CloseValue - OpenValue
                
                With ws.Cells(ResultIndex, 10)
                    .Value = YearlyChange
                    If YearlyChange >= 0 Then
                        .Interior.Color = RGB(0, 255, 0)
                    Else
                        .Interior.Color = RGB(255, 0, 0)
                    End If
                End With
                
                'Calculate, Save, and Format Percent Change
                PercentChange = YearlyChange / OpenValue
                With ws.Cells(ResultIndex, 11)
                    .Value = PercentChange
                    .NumberFormat = "0.00%"
                End With
                
                'Calculate and Save Total Volume
                TotalVolume = TotalVolume + ws.Cells(I, 7).Value
                ws.Cells(ResultIndex, 12).Value = TotalVolume
                
                'Increment to the next Ticker
                ResultIndex = ResultIndex + 1
                
                ' Reset values
                TotalVolume = 0
                IsFirstItem = True
                
            'Still in the same ticker
            Else
                'Get the opening Price from the beginning of the Year
                If IsFirstItem Then
                    IsFirstItem = False
                    OpenValue = ws.Cells(I, 3).Value
                End If
                TotalVolume = TotalVolume + ws.Cells(I, 7).Value
                
            End If
            
        Next I
        
        'Calculate "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
        Call Greatest(ws)

        'AutoFit Columns
        ws.Cells.EntireColumn.AutoFit
    Next ws


End Sub

'Calculate "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
Sub Greatest(ws):
    'Define Variables
    Dim Increase As Double
    Dim Decrease  As Double
    Dim Volume As LongLong
    Dim TickerIncrease As String
    Dim TickerDecrease  As String
    Dim TickerVolume  As String
    Dim KRows As Long
    
    'Find number of rows in column K
    KRows = ws.Cells(Rows.Count, "K").End(xlUp).Row
    
    ' Add the headers
    With ws
        .Cells(1, 16).Value = "Ticker"
        .Cells(1, 17).Value = "Value"
        .Cells(2, 15).Value = "Greatest % increase"
        .Cells(3, 15).Value = "Greatest % decrease"
        .Cells(4, 15).Value = "Greatest total volume"
    End With
    
    For I = 2 To KRows
        If ws.Cells(I, 11).Value > Increase Then
            Increase = ws.Cells(I, 11).Value
            TickerIncrease = ws.Cells(I, 9).Value
        End If
        If ws.Cells(I, 11).Value < Decrease Then
            Decrease = ws.Cells(I, 11).Value
            TickerDecrease = ws.Cells(I, 9).Value
        End If
        If ws.Cells(I, 12).Value > Volume Then
            Volume = ws.Cells(I, 12).Value
            TickerVolume = ws.Cells(I, 9).Value
        End If
    Next I
    
    With ws
        .Cells(2, 16).Value = TickerIncrease
        .Cells(2, 17).Value = Increase
        .Cells(2, 17).NumberFormat = "0.00%"
    
        .Cells(3, 16).Value = TickerDecrease
        .Cells(3, 17).Value = Decrease
        .Cells(3, 17).NumberFormat = "0.00%"
        
        .Cells(4, 16).Value = TickerVolume
        .Cells(4, 17).Value = Volume
    End With
    
End Sub


