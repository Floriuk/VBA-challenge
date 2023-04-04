Attribute VB_Name = "Module1"
Sub StockChange()
    For Each ws In Worksheets
        WorksheetName = ws.Name

        'Naming the column
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Cells(1, "Q").Value = "Ticker"
        ws.Cells(1, "R").Value = "Value"
        ws.Cells(2, "P").Value = "Greatest % Increase"
        ws.Cells(3, "P").Value = "Greatest % Decrease"
        ws.Cells(4, "P").Value = "Greatest Total Volume"
        
        'Setting the yearly change
        Dim Yearly_Change As Double
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Percentage_Change As Double
        Dim Brand_Counter As Long
        Dim Summary_Row2 As Integer
        
        Brand_Counter = 0
        Summary_Row2 = 2
        Open_price = 2
        Total = 0
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To LastRow
        
            'Check when the ticker symbol changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Total = Total + ws.Cells(i, "G").Value
                Yearly_Open = ws.Cells(Open_price, "C").Value
                Yearly_Close = ws.Cells(i, "F").Value
                Yearly_Change = Yearly_Close - Yearly_Open
                Percentage_Change = Yearly_Change / Yearly_Open
                
                'populating the ticker value and the yearly change
                ws.Range("I" & Summary_Row2).Value = ws.Cells(i, "A").Value
                ws.Range("J" & Summary_Row2).Value = Yearly_Change
                
                'Checking the value of the cell to fill it green for positive and red for negative
                If ws.Range("J" & Summary_Row2).Value > 0 Then
                    ws.Range("J" & Summary_Row2).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Summary_Row2).Interior.ColorIndex = 3
                End If
                
                'printing the percentage change and adding the percentage format
                ws.Range("K" & Summary_Row2).Value = Percentage_Change
                ws.Range("K" & Summary_Row2).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Row2).Value = Total
                
                Open_price = i + 1
                Summary_Row2 = Summary_Row2 + 1
                'setting the counter to 0 again
                Total = 0
            Else
                Total = Total + ws.Cells(i, "G").Value
                
            End If
        
        Next i

        'Determining the last row
        LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Attribute
        WorksheetName = ws.Name
        Dim Greatest_Percent As Double
        Dim Max_Ticker As Range
        Greatest_increase = 0
        Greatest_increase_ticker = ""
        Greatest_decrease = 0
        Greatest_decrease_ticker = ""
        Greatest_volume = 0
        Greatest_volume_ticker = ""
        
        For i = 2 To LastRow
            
            If ws.Cells(i, "K").Value > Greatest_increase Then
                Greatest_increase = ws.Cells(i, "K").Value
                Greatest_increase_ticker = ws.Cells(i, "I").Value
            End If
            
                
            If ws.Cells(i, "K").Value < Greatest_decrease Then
                Greatest_decrease = ws.Cells(i, "K").Value
                Greatest_decrease_ticker = ws.Cells(i, "I").Value
            End If
            
            If ws.Cells(i, "L").Value > Greatest_volume Then
                Greatest_volume = ws.Cells(i, "L").Value
                Greatest_volume_ticker = ws.Cells(i, "I").Value
            End If
            
        Next i
        
        'populating the results
        ws.Range("Q2").Value = Greatest_increase_ticker
        ws.Range("R2").Value = Greatest_increase
        ws.Range("R2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Greatest_decrease_ticker
        ws.Range("R3").Value = Greatest_decrease
        ws.Range("R3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Greatest_volume_ticker
        ws.Range("R4").Value = Greatest_volume
        
        ws.Columns("I:R").AutoFit
        
    
    Next ws
    MsgBox ("Complete")
    
End Sub

Sub Clearws()
    For Each ws In Worksheets
        WorksheetName = ws.Name
        ws.Columns("I:R").Clear
    Next ws
End Sub
