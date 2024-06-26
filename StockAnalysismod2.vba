
Sub alphabeticaltesting()

Dim rowindex As Long
Dim colindex As Integer
Dim lastrow As Long
Dim total As Double
Dim percentchange As Double
Dim change As Double
Dim ws As Worksheet
Dim start As Long
Dim dailychange As Single
Dim averagechange As Double
Dim days As Integer


For Each ws In Worksheets
    colindex = 0
    total = 0
    start = 2
    change = 0
    dailychange = 0
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "quarterlychange"
    ws.Cells(1, 11).Value = "percentchange"
    ws.Cells(1, 12).Value = "Totalstockvolume"
    ws.Range("O1").Value = "Tickerpercent"
    ws.Range("P1").Value = "Tickervalue"
    ws.Range("N2").Value = "Greatest%increase"
    ws.Range("N3").Value = "Greatest%decrease"
    ws.Range("N4").Value = "Greatesttotalvolume"
    
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For rowindex = 2 To lastrow
    
        
        If ws.Cells(rowindex + 1, 1) <> ws.Cells(rowindex, 1).Value Then
        
        total = total + ws.Cells(rowindex, 7).Value
        
          If total = 0 Then
          
            
            ws.Range("I" & 2 + colindex).Value = Cells(rowindex, 1).Value
            ws.Range("J" & 2 + colindex).Value = 0
            ws.Range("k" & 2 + colindex).Value = "%" & 0
            ws.Range("L" & 2 + colindex).Value = 0
            
            
          Else
            If ws.Cells(start, 3) = 0 Then
             For Value = start To rowindex
              If ws.Cells(Value, 3).Value <> 0 Then
                start = Value
                Exit For
              End If
              
             Next Value
             
             End If
             
             change = (ws.Cells(rowindex, 6) - ws.Cells(start, 3))
             percentchange = change / ws.Cells(start, 3)
             
             start = rowindex + 1
              
              ws.Range("I" & 2 + colindex).Value = ws.Cells(rowindex, 1).Value
              ws.Range("J" & 2 + colindex) = change
              ws.Range("J" & 2 + colindex).NumberFormat = "0.00"
              ws.Range("K" & 2 + colindex) = percentchange
              ws.Range("K" & 2 + colindex).NumberFormat = "0.00"
              
              ws.Range("L" & 2 + colindex).Value = total
              
              Select Case change
              Case Is > 0
                ws.Range("J" & 2 + colindex).Interior.ColorIndex = 4
              Case Is < 0
                ws.Range("J" & 2 + colindex).Interior.ColorIndex = 3
              Case Else
                ws.Range("J" & 2 + colindex).Interior.ColorIndex = 0
              End Select
              
           End If
           
           total = 0
           change = 0
           colindex = colindex + 1
           days = 0
           dailychange = 0
           
         Else
          total = total + ws.Cells(rowindex, 7).Value
                              
          End If
           
    Next rowindex
    
    ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("P3") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("P4") = "%" & WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    
    ws.Range("O2") = ws.Cells(increase_number + 1, 9)
    ws.Range("O3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("O4") = ws.Cells(volume_number + 1, 9)
    
    Next ws


End Sub







