Attribute VB_Name = "Module1"
Sub VBA_Of_Wall_Street()

    'Set Start Variables
    'Set First Stage Variables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim Total_Stock_Volume As Double
        
        'Set Second Stage Variables
        Dim Open_Stock_Price As Double
        Dim Close_Stock_Price As Double
        Dim Open_Stock_PriceRow As Long

            'Reset Stock Volume
            Total_Stock_Volume = 0

                 'Set Variable for summary table
                 Dim Summary_Table_Row As Integer
                     
                     'Loop through all worksheets in workbook
                    For Each ws In Worksheets
                    
                        'Challenge Headers for Columns
                        Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Value")
                        ws.Range("P1").Value = "Ticker"
                        ws.Range("Q1").Value = "Value"
                        ws.Range("O2").Value = "Greatest % Increase"
                        ws.Range("O3").Value = "Greatest % Decrease"
                        ws.Range("O4").Value = "Greatest Total Volume"

                            'Location Start Of  Summary Table
                             Summary_Table_Row = 2
    
                                 'Set Row Count
                                 Open_Stock_PriceRow = 2
                                 Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
                                 
                                    'Begin Loop
                                    For i = 2 To Last_Row
                                         On Error Resume Next
                                            'Check Ticker Value is same.....if not
                                             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                             Ticker = ws.Cells(i, 1).Value
            
                                                'Values for Yearly Change Calculation
                                                Open_Stock_Price = ws.Cells(Open_Stock_PriceRow, 3).Value
                                                Close_Stock_Price = ws.Cells(i, 6).Value
                                                    'Input Value of Change
                                                    Yearly_Change = Close_Stock_Price - Open_Stock_Price
                                                         'Reset Value
                                                        If Open_Stock_Price = 0 Then
                                                        '   Reset Value
                                                            Percentage_Change = 0
        Else
                                                                'Print Value of Percentage Difference
                                                                Percentage_Change = Yearly_Change / Open_Stock_Price
        End If
                                                                    'Print Value of Total Stock
                                                                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
                                                                'Summary Table
                                                                ws.Range("I" & Summary_Table_Row).Value = Ticker
                                                                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                                                        'Summary Table Colour Value
                                                        If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                                                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
                                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
                                                ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
                                                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                                        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                        Summary_Table_Row = Summary_Table_Row + 1
                                'Reset Stock Annual Values before loop
                                Open_Stock_PriceRow = i + 1
                                Total_Stock_Volume = 0
        Else              'Print Stock Volume Values
                            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If
        Next i
                        'Set Location for Value input of Summary Row
                        lRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
                        minValue = 0
                        maxValue = 0
                        maxTotal_Stock_Volume = 0
                        
                             'Set Start of Challenge Loop
                            For i = 2 To lRow
                        
                                    'Greatest % Increase Location
                                    If ws.Cells(i, 11) > maxValue Then
                                    maxValue = ws.Cells(i, 11)
                                    maxTicker = ws.Cells(i, 9)
        Else                           'Print Value
                                         maxValue = maxValue
                                         
                                            
        End If                          'Greatest % Decrease Location
                                            If ws.Cells(i, 11) < minValue Then
                                            minValue = ws.Cells(i, 11)
                                            minTicker = ws.Cells(i, 9)
        Else                              'Print Value
                                            minValue = minValue
        End If
                                                    'Greatest Stock Volume Locatiion
                                                     If ws.Cells(i, 12) > maxTotal_Stock_Volume Then
                                                     maxTotal_Stock_Volume = ws.Cells(i, 12)
                                                     maxTotal_Stock_VolumeTicker = ws.Cells(i, 9)
        Else                                      'Print Value
                                                    maxTotal_Stock_Volume = maxTotal_Stock_Volume
        End If
        Next i
                                     'Challenge Summary Values
                                     ws.Range("P2").Value = maxTicker
                                     ws.Range("Q2").Value = maxValue
                                     ws.Range("Q2").NumberFormat = "0.00%"
                                     ws.Range("P3").Value = minTicker
                                     ws.Range("Q3").Value = minValue
                                     ws.Range("Q3").NumberFormat = "0.00%"
                                     ws.Range("P4").Value = maxTotal_Stock_VolumeTicker
                                     ws.Range("Q4").Value = maxTotal_Stock_Volume
    
                            'Resize Columns
                            ws.Columns("I:Q").AutoFit
    
        Next ws

        End Sub
