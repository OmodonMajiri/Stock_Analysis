Attribute VB_Name = "Module1"
Sub StockData():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        ' Row
        Dim i As Long
        ' Start row of ticker
        Dim j As Long
        ' Index counter to fill Ticker row
        Dim TickCount As Long
        ' Last row column A
        Dim LastRowA As Long
        ' Last row column I
        Dim LastRowI As Long
        ' Variable for percent change calculation
        Dim GreatIncr As Double
        ' Variable for greatest increase ticker
        Dim GreatIncr_Ticker As String
        ' Variable for greatest decreas calculation
        Dim GreatDecr As Double
        ' Variable for greatest decreas ticker
        Dim GreatDecr_Ticker As String
        ' Variable for greatest total volume
        Dim GreatVol As Double
        ' Variable for greatest total volume ticker
        Dim GreatVol_Ticker As String
        
        ' Get Worksheet Name
        WorksheetName = ws.Name
        
        ' Create Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Set Ticker Counter to first row
        TickCount = 2
        
        ' Set start row to 2
        j = 2
        

        ' Find the last non-blank cell in Column A
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox ("Last row in column A is " & LostRowA)
        
    
            ' Loop through all the rows
            For i = 2 To LastRowA
        
                ' check to see if ticker name has changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' write ticker in column I
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
            
                ' calculate and write year change in column J
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
            
                    ' Conditional formatin
                    If ws.Cells(TickCount, 10).Value > 0 Then
                
                    ' Set cell color to green
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                
                    Else
                
                    ' Set cell color to red
                    ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                
                    End If
                
                ' Calculate and write percentage change in column K
                If ws.Cells(j, 3).Value <> 0 Then
                PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                ' % format
                ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                
                Else
                
                ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                
                End If
                
            ' Calculate and write total volumn in column L
            ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
            ' Increase the TickCount by 1
            TickCount = TickCount + 1
            
            ' Set the new start row of ticker
            j = i + 1
            
            End If
            
        Next i
            
      ' Find last non-empty cell in column I
      LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
      'MsgBox ("Last row in culumn I is " & LastRowI)
      
      ' Summary
      GreatVol = ws.Cells(2, 12).Value
      GreatIncr = ws.Cells(2, 11).Value
      GreatDecr = ws.Cells(2, 11).Value
      
        ' Loop for summary
        GreatVol = 0
        For i = 2 To LastRowI
        
            ' Greatest Total Volume
            If ws.Cells(i, 12).Value > GreatVol Then
              GreatVol = ws.Cells(i, 12).Value
              GreatVol_Ticker = ws.Cells(i, 9).Value
            End If
            
            ' Greatest Increase
            If ws.Cells(i, 11).Value > GreatIncr Then
              GreatIncr = ws.Cells(i, 11).Value
              GreatIncr_Ticker = ws.Cells(i, 9).Value
            End If
            
            ' Greatest Decrease
            If ws.Cells(i, 11).Value < GreatDecr Then
              GreatDecr = ws.Cells(i, 11).Value
              GreatDecr_Ticker = ws.Cells(i, 9).Value
            End If
            
            
            
        ' Write result in ws.cells
        ws.Cells(2, 16).Value = GreatIncr_Ticker
        ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
        ws.Cells(3, 16).Value = GreatIncr_Ticker
        ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
        ws.Cells(4, 16).Value = GreatVol_Ticker
        ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
        Next i
        
    ' Adjust column width
    
    Worksheets(WorksheetName).Columns("A:Z").AutoFit
    
        
  Next ws
                
        
        
    
End Sub
