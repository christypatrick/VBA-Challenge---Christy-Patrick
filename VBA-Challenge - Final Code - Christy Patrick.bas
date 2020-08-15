Attribute VB_Name = "Module1"
Sub StockTicker():
       
    ' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
    
        ' Set an initial variable for holding the Ticker name
        Dim Ticker As String
           
        'Set variables for opening and closing dates
        Dim Opening_Date As Double
        Dim Closing_Date As Double
        
        'Set variables for Beginning and End of Year
        Dim Beginning_Year As Range
        Dim End_Year As Range
        
        'Set an initial variable for holding opening amount
        Dim Opening_Amount As Double
        Opening_Amount = 0
        
        'Set an initial variable for holding closing amount
        Dim Closing_Amount As Double
        Closing_Amount = 0
        
        'Set initial variable for holding Yearly Change
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        'Set initial variable for Percent Change
        Dim Percent_Change As Double
        Percent_Change = 0
        
        ' Set an initial variable for holding the total per credit card brand
        Dim Ticker_Volume As Double
        Ticker_Volume = 0
        
        'Set an initial variable for the Greatest Percent Change Increase
        Dim Greatest_Increase As Double
        Greatest_Increase = 0
        
        'Set an initial variable for the Greatest Percent Change Decrease
        Dim Greatest_Decrease As Double
        Greatest_Decrease = 0
        
        'Set an initial variable for the Greatest Volume
        Dim Greatest_Volume As Double
        Greatest_Volume = 0
           
        ' Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
          
            'Set Column Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
                    
            'Identify Last Row in the Worksheet
            Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
               
        ' Loop through all ticker entries to calculate Volume
          For i = 2 To Lastrow
                   
              ' Check if we are still within the same ticker, if it is not...
              If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                ' Set the Ticker name
                Ticker = ws.Cells(i, 1).Value
                          
                'Set Closing Amount
                Closing_Amount = ws.Cells(i, 6)
                
                'Calculate Yearly Change
                Yearly_Change = Closing_Amount - Opening_Amount
                         
                ' Print the Yearly Change to the Summary Table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                     'If Yearly_Change is greater than 0, make it green.  Otherwise make it red.
                    If Yearly_Change >= 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        
                    End If
   
                'Calculate Percent Change, sets value for the current ticker
                    If Opening_Amount > 0 And Closing_Amount > 0 Then
                        Percent_Change = Yearly_Change / Opening_Amount
                    
                    Else
                        Percent_Change = 0
                    
                    End If
                                                                                                       
                'Find Greatest Percent Change Increase (Is current ticker % greater than previous ticker%)
                    If Percent_Change > Greatest_Increase Then
                        Greatest_Increase = Percent_Change
                        ws.Range("O2") = Ticker
                        ws.Range("P2") = Format((Greatest_Increase), "Percent")

                    End If
                    
                'Find Greatest Percent Change Decrease (Is current ticker % greater than previous ticker%)
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    ws.Range("O3") = Ticker
                    ws.Range("P3") = Format((Greatest_Decrease), "Percent")

                End If

                ' Print the Percent Change to the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Format((Percent_Change), "Percent")
                
                ' Add to the Ticker Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                'Find Greatest Total Volume (Is current ticker volume greater than previous ticker volume)
                If Ticker_Volume > Greatest_Volume Then
                    Greatest_Volume = Ticker_Volume
                    ws.Range("O4") = Ticker
                    ws.Range("P4") = Greatest_Volume

                End If
                
                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                
                ' Print the Ticker Volume to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Volume
          
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                                                          
                ' Reset the Ticker Volume
                Ticker_Volume = 0
    
                ' Check if we are still within the same ticker, if it is not...
              
           ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
         
                Opening_Amount = ws.Cells(i, 3).Value
                
                 ' Add to the Ticker Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
            
             ' If the cell immediately following a row is the same ticker...
            Else
    
                ' Add to the Ticker Volume
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                                                                                              
           
            End If
        
         Next i
        
            ' Autofit to display data
            ws.Columns("I:P").AutoFit
        
   'Next ws
    Next ws
        
End Sub
