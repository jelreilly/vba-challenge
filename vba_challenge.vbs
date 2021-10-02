Sub Wall_Street()


'Set an initial value for holding ticker
Dim Ticker As String

'Set an initial variable for holding the total per ticker
Dim Total_Volume As Double
Total_Volume = 0
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0
Dim i As Long
Dim lRow As Long

'New Method For Open
Dim Z As Long
Z = 2

'Title summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"




'Keep track of the locaion for each ticker in the summary table
Dim Summary_Table_Row As String
Summary_Table_Row = 2


        'Loop through all the tickers
        
     lRow = Cells(Rows.Count, 1).End(xlUp).Row
     
    For i = 2 To lRow
            
            'Check if we are still in the same ticker, if it is not
           If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            
             
                'Set the ticker name
                 Ticker = Cells(i, 1).Value
             
                'Add to the volume total
                 Total_Volume = Total_Volume + Cells(i, 7).Value
                 
                 'Open Price Checker
                  If Cells(Z, 3).Value = 0 Then
                    For FindValue = Z To i
                        If ws.Cells(FindValue, 3).Value <> 0 Then
                            Z = FindValue
                            Exit For
                        End If
                    Next FindValue
                   End If
                
                            
                'closing price from Cells(i, 6).Value
                 Close_Price = Cells(i, 6).Value
                 
                'Print the Ticker in the summary table
                 Range("I" & Summary_Table_Row).Value = Ticker
             
                'Print the total volume in the summary table
                 Range("L" & Summary_Table_Row).Value = Total_Volume
                 
                          
                  ' calculate yearly change where cells (i,6) is close cells (z,3) is open
                  Yearly_Change = Cells(i, 6).Value - Cells(Z, 3).Value

                    
                    ' calculate percentage change
                  Percent_Change = (Cells(i, 6).Value - Cells(Z, 3).Value) / Cells(Z, 3).Value
                    
                    
                    'Print the Percent Change in summary table
                   Range("K" & Summary_Table_Row).Value = Percent_Change
                     
                     'Print the Yearly Change in the Summary Table
                   Range("J" & Summary_Table_Row).Value = Yearly_Change
                   
                   'Add 1 to the summary table row
                   Summary_Table_Row = Summary_Table_Row + 1
                   
                   'Reset the Total Volume
                   Total_Volume = 0
                
                  Z = i + 1
                  
                Else
                    
                    'Add to the Total Volume
                    Total_Volume = Total_Volume + Cells(i, 7).Value
                    
                    
                    
                                       
                End If
                    
                    'Add format to percent change
                    Cells(i, 11).NumberFormat = "0.00%"
                    
                    
                    'Add format to Yearly change
                If Cells(i, 10).Value >= 0 Then
                    
                    Cells(i, 10).Interior.ColorIndex = 4
                    
                Else
                    Cells(i, 10).Interior.ColorIndex = 3
                    
                End If
                
                        
                 
            
                 
        
                                   
       
       
            Next i
    



End Sub
