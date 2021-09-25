Attribute VB_Name = "Module1"
Sub vbaProj():

    'loop through each worksheet
    For Each ws In Worksheets
        
        'get the row length of the worksheet we're in
        rowL = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'summary table header
        ws.Range("I" & 1).Value = "Ticker"
        ws.Range("J" & 1).Value = "Yearly Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock Volume"
        
        'summary row counter
        sumR = 2
        
        'Ticker Total Stock Volume
        Dim totStockV As Double
        'opening price
        Dim openP As Double
        'closing price
        Dim closeP As Double
        'yearly change for ease of calculation
        Dim yrlyChange As Double
        
        Dim greatPIV As Double 'greatest percent increase
        Dim greatPDV As Double 'greatest percent decrease
        Dim greatTotVV As Double 'greatest total volume
        
        Dim greatPIT As String 'greatest percent increase Ticker
        Dim greatPDT As String 'greatest percent decrease Ticker
        Dim greatTotVT As String 'greatest total volume Ticker
        
        'second summary table headers and row lables
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("Q" & 1).Value = "Value"
        
        
        'loop through each row in worksheet
        For r = 2 To rowL
            'checking to see if the row i'm on is same as the next row. If the first If is true, I'm on the last row of the same ticker
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                
                totStockV = totStockV + ws.Cells(r, 7).Value 'add the volume to the total
                closeP = ws.Cells(r, 6).Value
                
                ws.Range("I" & sumR).Value = ws.Cells(r, 1).Value
                yrlyChange = closeP - openP
                ws.Range("J" & sumR).Value = yrlyChange 'Yearly Change
                ws.Range("J" & sumR).NumberFormat = "#,##0.00" 'format Yearly Change
                ws.Range("L" & sumR).Value = totStockV 'Total Stock Volume
                
                If openP <> 0 Then 'make sure you are not dividing by 0
                    ws.Range("K" & sumR).Value = yrlyChange / openP 'Percent Change
                    ws.Range("K" & sumR).NumberFormat = "0.00%" 'format percent
                    
                    'check for greatest percent increase
                    If yrlyChange / openP > greatPIV Then
                        greatPIV = yrlyChange / openP
                        greatPIT = ws.Cells(r, 1).Value
                    End If
                    
                    'check for greatest percent decrease
                    If yrlyChange / openP < greatPDV Then
                        greatPDV = yrlyChange / openP
                        greatPDT = ws.Cells(r, 1).Value
                    End If
                Else
                    ws.Range("K" & sumR).Value = 0 'Percent Change
                    ws.Range("K" & sumR).NumberFormat = "0.00%" 'format percent
                End If
                
                'conditional for changing color
                If (closeP - openP) >= 0 Then
                    ws.Range("J" & sumR).Interior.ColorIndex = 10 'green if it's positive
                Else
                    ws.Range("J" & sumR).Interior.ColorIndex = 3 'red if it's negative
                End If
                
                
                'check for greatest Total Volume
                If totStockV > greatTotVV Then
                    greatTotVV = totStockV
                    greatTotVT = ws.Cells(r, 1).Value
                End If
                
                'moving down in the summary table
                sumR = sumR + 1
                
                'reset total stock volume
                totStockV = 0
                
                
            Else
                totStockV = totStockV + ws.Cells(r, 7).Value 'add the volume to the total
                
                If ws.Cells(r, 1).Value <> ws.Cells(r - 1, 1).Value Then 'check to see if it's first row
                    openP = ws.Cells(r, 3).Value 'store value to calculate percent change later
                End If
                
            End If
        
        Next r
        
        ws.Range("Q" & 2).Value = greatPIV 'Greatest % Increase
        ws.Range("Q" & 2).NumberFormat = "0.00%" 'format percent
        ws.Range("Q" & 3).Value = greatPDV 'Greatest % Decrease
        ws.Range("Q" & 3).NumberFormat = "0.00%" 'format percent
        ws.Range("Q" & 4).Value = greatTotVV 'Greatest Total Volume
        ws.Range("P" & 2).Value = greatPIT 'Greatest % Increase Ticker
        ws.Range("P" & 3).Value = greatPDT 'Greatest % Decrease Ticker
        ws.Range("P" & 4).Value = greatTotVT 'Greatest Total Volume Ticker
               
        greatPIV = 0
        greatPDV = 0
        greatTotVV = 0
        
        greatPIT = ""
        greatPDT = ""
        greatTotVT = ""
    
    Next ws

End Sub

