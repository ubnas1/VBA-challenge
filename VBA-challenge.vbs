Sub stock_analysis()

    '------------------------------
    ' Loop through all worksheets
    '------------------------------
    For Each ws In Worksheets

        
        'create a variable to hold sheet name
        Dim sheet_name As String
        
        'grab the worksheet name
        sheet_name = ws.Name
        
        'Determine the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set initial variable for holding ticker name
        Dim ticker As String
        
        ' Set initial variable for holding stock volume
        Dim stock_vol As Double
        
        ' Set initial variable for holding yearly change
        Dim year_change As Double
        
        ' Set initial variable for holding percent change
        Dim per_change As Double
        
        ' Set initial variable for holding open price
        Dim open_price As Double
        
        'set open price to 0
        open_price = 0
        
        ' Set initial variable for greatest % increase
        Dim g_per_increase As Double
        g_per_increase = 0
        
        ' Set initial variable for greatest % decrease
        Dim g_per_decrease As Double
        g_per_decrease = 0
        
        ' Set initial variable for greatest total volume
        Dim g_tot_volume As Double
        g_tot_volume = 0
        
        ' set initial variable for row count
        Dim sum_row As Integer
        sum_row = 2
        
        
        ' Set initial variable for holding close price
        Dim close_price As Double
        close_price = 0
        
        
        ' Add New columns for headers and do their formatting for improved visibility
        ws.Range("I1").EntireColumn.Insert
        ws.Range("I1").Value = "Ticker"
        ws.Range("I1").Interior.ColorIndex = 19
        
        ws.Range("J1").EntireColumn.Insert
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("J1").Interior.ColorIndex = 19
        
        ws.Range("K1").EntireColumn.Insert
        ws.Range("K1").Value = "% Change"
        ws.Range("K1").Interior.ColorIndex = 19
            
        ws.Range("L1").EntireColumn.Insert
        ws.Range("L1").Value = "Total Stock Vol."
        ws.Range("L1").Interior.ColorIndex = 19
        
        ws.Range("M1").EntireColumn.Insert
        ws.Range("M1").Value = "Open Price"
        ws.Range("M1").Interior.ColorIndex = 19
        
        ws.Range("N1").EntireColumn.Insert
        ws.Range("N1").Value = "Close Price"
        ws.Range("N1").Interior.ColorIndex = 19
    
        ws.Range("R1").Value = "Ticker"
        ws.Range("R1").Interior.ColorIndex = 19
        
        ws.Range("S1").Value = "Value"
        ws.Range("S1").Interior.ColorIndex = 19
        
        ws.Range("P2").Value = "Greatest % increase"
        ws.Range("P2:Q2").Interior.ColorIndex = 35
        
        ws.Range("P3").Value = "Greatest % decrease"
        ws.Range("P3:Q3").Interior.ColorIndex = 35
        
        ws.Range("P4").Value = "Greatest Total Volume"
        ws.Range("P4:Q4").Interior.ColorIndex = 35
        
        ' start of loop to loop through every piece of data in all sheets
        For i = 2 To last_row
        
            'check if we are in the same ticker, if not then enter this loop
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'grab ticker name
                ticker = ws.Cells(i, 1).Value
                        
                ' grab the close price
                close_price = ws.Cells(i, 6).Value
                
                'print the close price
                ws.Range("N" & sum_row).Value = close_price
                
                'add to stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                
                'print total stock volume to the summary table
                ws.Range("L" & sum_row).Value = stock_vol
                
                ' print ticker
                ws.Range("I" & sum_row).Value = ticker
                
                'calculate year change
                year_change = close_price - open_price
                
                'conditional formatting for yearly change
                If year_change < 0 Then
                
                    'Red color cell for negative year change value
                    ws.Range("J" & sum_row).Interior.ColorIndex = 3
                Else
                    'Green color cell for 0 or above yearly change value
                    ws.Range("J" & sum_row).Interior.ColorIndex = 4
                End If

                'print year change
                ws.Range("J" & sum_row).Value = year_change
                            
                'grab % change
                '-------------
                
                'catch the divisible by 0 runtime error
                If open_price = 0 Then
                    per_change = close_price
                Else
                    per_change = ((close_price - open_price) / open_price) * 100
                
                End If
                
                'print % change
                ws.Range("K" & sum_row).Value = per_change
                
                'conditional formatting for % change
                If per_change < 0 Then
                    
                    'Red color cell for negative % change value
                    ws.Range("K" & sum_row).Interior.ColorIndex = 3 'Red
                Else
                    'Green color cell for positive % change value
                    ws.Range("K" & sum_row).Interior.ColorIndex = 4 'Green
                End If
                
                
                'get greatest total stock value
                If stock_vol > g_tot_volume Then
                    g_tot_volume = stock_vol
                
                    ws.Range("S4").Value = g_tot_volume
                    ws.Range("R4").Value = ticker
                
                End If
                
                
                'get greatest % change increase value
                If per_change > g_per_increase Then
                    g_per_increase = per_change
                    
                    ws.Range("R2").Value = ticker
                    ws.Range("S2").Value = g_per_increase
                
                End If
                
                
                
                'get greatest % change decrease value
                If per_change < g_per_decrease Then
                    g_per_decrease = per_change
                    
                    ws.Range("R3").Value = ticker
                    ws.Range("S3").Value = g_per_decrease
                
                
                End If
                
                'go to next row count
                sum_row = sum_row + 1
                
                
                'reset the stock volume, open price, close price
                stock_vol = 0
                close_price = 0
                open_price = 0
                
            Else
            
                ' if we are in the same stock (ticker)
                'add to stock volume
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                    
                'If new ticker
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    
                    'grab the open price value
                    open_price = ws.Cells(i, 3).Value
                    ws.Range("M" & sum_row).Value = open_price
                End If
            
            End If
        
        Next i
        
    Next ws

End Sub


