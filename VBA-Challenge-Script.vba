Attribute VB_Name = "Module1"
Sub Stock_Metrics()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Quarter1")
Dim Lastrow As Long
Dim Ticker As String
Dim openingprice As Double
Dim closingprice As Double
Dim totalvolume As Double
Dim currentquarter As String
Dim i As Long
Dim outputrow As Long
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double
Dim stockincrease As String
Dim stockdecrease As String
Dim stockvolume As String


Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

greatestincrease = -1 * (WorksheetFunction.Max(ws.Range("K2:K" & Lastrow)) - 1)
greatestdecrease = WorksheetFunction.Min(ws.Range("K2:K" & Lastrow)) + 1
greatestvolume = -1



For Each ws In Worksheets

    outputrow = 1
    For i = 2 To Lastrow
        Ticker = ws.Cells(i, 1).Value
    
         If i = 2 Or ws.Cells(i - 1, 1).Value <> Ticker Or quarter <> currentquarter Then
              openingprice = ws.Cells(i, 3).Value
              closingprice = ws.Cells(i, 6).Value
              totalvolume = ws.Cells(i, 7).Value
              currentquarter = quarter
        Else
              closingprice = ws.Cells(i, 6).Value
              totalvolume = totalvolume + ws.Cells(i, 7).Value
        End If
    
    If i = Lastrow Or ws.Cells(i + 1, 1).Value <> Ticker Then
        Dim change As Double
        Dim percentagechange As Double
        change = closingprice - openingprice
            If openingprice <> 0 Then
                 percentagechange = Round((change / openingprice * 100), 2)
                 'ws.Cells(i, 11).NumberFormat = "0.00"
            Else
                percentagechange = 0
            End If
        
                ws.Cells(outputrow + 1, 9).Value = Ticker
                ws.Cells(outputrow + 1, 10).Value = change
                'ws.Cells(outputrow + 1, 11).Value = Format(percentagechange, "0.00") & "%"
                'ws.Cells(outputrow + 1, 11).Value = Format(percentagechange, "#.##") & "%"
                ws.Cells(outputrow + 1, 11).Value = percentagechange
                ws.Cells(outputrow + 1, 12).Value = totalvolume
        
                If change < 0 Then
                    ws.Cells(outputrow + 1, 10).Interior.ColorIndex = 3
                ElseIf change > 0 Then
                    ws.Cells(outputrow + 1, 10).Interior.ColorIndex = 4
                End If
                                      
            If percentagechange > greatestincrease Then
                greatestincrease = percentagechange
                stockincrease = Ticker
            End If
            
            If percentagechange < greatestdecrease Then
                greatestdecrease = percentagechange
                stockdecrease = Ticker
            End If
            
            If totalvolume > greatestvolume Then
                greatestvolume = totalvolume
                stockvolume = Ticker
            End If
    
    
            outputrow = outputrow + 1
        End If
    
      Next i
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Volume"
            
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Incease"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume:"
            
            ws.Cells(2, 16).Value = stockincrease
            ws.Cells(3, 16).Value = stockdecrease
            ws.Cells(4, 16).Value = stockvolume
            
            ws.Cells(2, 17).Value = greatestincrease
            ws.Cells(3, 17).Value = greatestdecrease
            ws.Cells(4, 17).Value = greatestvolume
                                         
Next ws

End Sub

