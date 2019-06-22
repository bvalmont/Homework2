Attribute VB_Name = "Homework2"
Sub Stockticker()
        
    
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    ws.Activate
    
    
 
            'Set variale for name of stock
        Dim Ticker As String
                'Set variable to hold total volume per stock
      Dim Total_Stock_Volume As Currency
      
      Total_Stock_Volume = 0
                'Track location for each stock in summary table
      Dim Summary_Table_Row As Integer
            
      
      Summary_Table_Row = 2
                'Loop trough all volumes
        For I = 2 To 797712
            'This is a way to check if its the same stock or not
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                'Set the stock name
            Ticker = Cells(I, 1).Value
                    'Add to stock total
                Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
                    'Print stock name in summary table
                Range("I" & Summary_Table_Row).Value = Ticker
                    'Print volumes in summary table
                Range("J" & Summary_Table_Row).Value = Total_Stock_Volume
                  
                        'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                    'reset stock volume
                Total_Stock_Volume = 0
                'if  the cell following a row is the same stock
                  Else
                'Add to the volume total
                    Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
                    
                       End If
                        
                        
                        Cells(I, 10).NumberFormat = "0"
                        
                        
                        
    Next I
        
                    Range("I1").Value = "Ticker"
                    
                    Range("J1").Value = "Total Volume"
            
            
   Next
   
    
    
    
            
    End Sub
    
    
    
    
        
         
      
      
      
