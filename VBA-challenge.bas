Attribute VB_Name = "Module1"
Sub Stock_Ticker()
    
    ' Print headers in appropriate cells
    Range("I1") = ("ticker")
    Range("J1") = ("yearly change")
    Range("K1") = ("percent change")
    Range("L1") = ("total stock volume")
    Range("M1") = ("open stock")
    Range("N1") = ("close stock")

    ' Set an initial variable for holding the stock ticker name
    Dim Stock_Ticker As String
    
    ' Set an initial variable for holding the total stock volume
    Dim Stock_Volume As Single
    Stock_Volume = 0
    
    'Set an initial variable for opening stock amount
    Dim Open_Stock As Double
    Open_Stock = 0
    
    ' Set an initial variable for closing stock amount
    Dim Close_Stock As Double
    Close_Stock = 0
    
    ' Set an initial variable for yearly change stock amount
    Dim Yearly_Change As Double
    Yearly_Change = 0
    
    'Set an initial variable for percentage change
    Dim Percent_Change As Single
    Percent_Change = 0
    
    'Set variable for last row
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Keep track of location of location of data
    Dim Summary_Stock_Table_Row As Double
    Summary_Stock_Table_Row = 2

    
    ' Loop through all stock volumes
    For i = 2 To Last_Row
       
                     
        ' Check if still in same ticker symbol, if not then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the open stock
            Open_Stock = Cells(i, 3).Value
            Range("M" & Summary_Stock_Table_Row).Value = Open_Stock
            
            ' Set the stock ticker
            Stock_Ticker = Cells(i, 1).Value
        
            'Add to the stock volume total
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
            ' Set the close stock amount
            Close_Stock = Cells(i, 6).Value
            Range("N" & Summary_Stock_Table_Row).Value = Close_Stock
            
            'Print the Stock Ticker in the in the Summary Column ticker
            Range("I" & Summary_Stock_Table_Row).Value = Stock_Ticker
        
            'Print the stock volume total in the Summary Column volume
            Range("L" & Summary_Stock_Table_Row).Value = Stock_Volume
           
           
            
            'Print the yearly change in the Summary Column yearly change
            Yearly_Change = Close_Stock - Open_Stock

            Range("J" & Summary_Stock_Table_Row).Value = Yearly_Change
            
            'Print the percent change in the Summary Column
        
            Percent_Change = Yearly_Change / Open_Stock * 100
            Range("K" & Summary_Stock_Table_Row).Value = Percent_Change
        
               
           ' Reset the Yearly Change Total
            Yearly_Change = 0
             
        
            ' Add one to the Summary Table Column Row
            Summary_Stock_Table_Row = Summary_Stock_Table_Row + 1
        
            ' Reset the Stock Volume Total
            Stock_Volume = 0
            
            
        'If the cell immediately following a row is the same stock ticker..
        Else
        
            'Add to the Stock Volume Total
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
        
    
       
       
        End If
        
      Next i
     
    

End Sub

