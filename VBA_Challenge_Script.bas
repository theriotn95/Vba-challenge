
Sub Stock_Data_2014__2016()

Dim ws As Worksheet

For Each ws In Worksheets

    'create summary table headers'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    Dim ticker As String
    
    'where summary table starts'
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
  'set variables for opening  and closing values'
    Dim openPrice As Double
    Dim openprice_ind As Double
    Dim closePrice As Double
    Dim yearly_change As Double
    Dim ticker_volume As Double
    Dim percent_change As Double
   
    openprice_ind = 2
    ticker_volume = 0
    
    
    'set your last row & loop through tickers'
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        openPrice = ws.Cells(openprice_ind, 3).Value
        
        'make sure still within same ticker'
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'set ticker'
        ticker = ws.Cells(i, 1).Value
        
        'add to volume total'
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
         
        'find year end price'
        closePrice = ws.Cells(i, 6).Value
            
        'calculate yearly change '
            yearly_change = closePrice - openPrice
            ws.Cells(i, 10).Value = yearly_change
                    
            
            'calculate percent_change'
            If openPrice = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / openPrice
            End If
            
            'Display values in the Summary Table'
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            'make the format percentage'
            ws.Range("K" & Summary_Table_Row) = Format(percent_change, "Percent")
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("L" & Summary_Table_Row).Value = ticker_volume

            'color format yearly change'
            If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
           
           'Add one to the summary table row'
          Summary_Table_Row = Summary_Table_Row + 1
          
          'Reset ticker volume'
          
          ticker_volume = 0
          
          'to make sure it has the correct #'
          openprice_ind = (i + 1)
          
            Else
        
        'make sure to add to ticker volume'
        ticker_volume = ticker_volume + ws.Cells(i, 7).Value
        
            End If
        
        Next i
        
        'ensure everythings fits'
        ws.Range("A:M").Columns.AutoFit
        
        
        Next ws
        
        
                      
                
                End Sub
    
