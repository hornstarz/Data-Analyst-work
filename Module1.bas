Attribute VB_Name = "Module1"
Sub stock_analysis()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim total As Double
    Dim change As Double
    Dim i As Long
    Dim j As Long
    Dim start As Long
    Dim rowCount As Long
    Dim greatestIncrease As Double
    Dim greatestTicker As String
    Dim greatestDecrease As Double
    Dim worstTicker As String
    Dim greatestVolume As Double
    Dim bestVolumeticker As String
    For Each ws In ThisWorkbook.Worksheets
    
    
      
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    j = 0
    total = 0
    change = 0
    start = 2

    rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    For i = 2 To rowCount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total = total + ws.Cells(i, 7).Value
            
            If total = 0 Then
                ws.Cells(2 + j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(2 + j, 10).Value = 0
                ws.Cells(2 + j, 11).Value = "%" & 0
                ws.Cells(2 + j, 12).Value = 0
            Else
                
                If ws.Cells(start, 3).Value = 0 Then
                    Dim find_value As Long
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
                If ws.Cells(start, 3).Value <> 0 Then
                    PercentChange = change / ws.Cells(start, 3).Value
                Else
                    PercentChange = 0
                End If
                
                ws.Cells(2 + j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(2 + j, 10).Value = change
                ws.Cells(2 + j, 11).Value = PercentChange
                ws.Cells(2 + j, 11).NumberFormat = "0.00%"
                ws.Cells(2 + j, 12).Value = total
                
                If PercentChange > greatestIncrease Then
                    greatestIncrease = PercentChange
                    greatestTicker = ws.Cells(i, 1).Value
                End If
            
                If PercentChange < greatestDecrease Then
                    greatestDecrease = PercentChange
                    worstTicker = ws.Cells(i, 1).Value
                End If
                              
                If total > greatestVolume Then
                    greatestVolume = total
                    bestVolumeticker = ws.Cells(i, 1).Value
                End If
                
                
            End If

            total = 0
            change = 0
            j = j + 1
            start = i + 1
        Else
            total = total + ws.Cells(i, 7).Value
        End If
    Next i
    
    ws.Cells(2, 16).Value = greatestTicker
    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = worstTicker
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = bestVolumeticker
    ws.Cells(4, 17).Value = greatestVolume
    ws.Columns.AutoFit
    With ws.Range("J2:J" & lastRow)
                    .FormatConditions.Delete
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:=0
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 255, 0)
                    .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:=0
                    .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)
                End With
    Next ws
End Sub
