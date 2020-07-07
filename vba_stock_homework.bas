Attribute VB_Name = "Module2"

Sub tickertron()

    '' There was supposed to be other stuff I did with these variables, but some of them just were forgot
    
Dim open_price, close_price, yearly_change As Double
Dim pct_change As Double
Dim biggest_pct_decrease As Double
Dim biggest_pct_increase As Double
Dim WS As Variant
Dim biggest_volume As Long
Dim ticker As String
Dim volume As Variant
Dim summary_row As Integer
Dim rowcount As Variant
Dim increase_row As Long
Dim decrease_row As Long
Dim increase_ticker As String
Dim decrease_ticker As String


summary_row = 2
rowcount = Cells(Rows.Count, 1).End(xlUp).Row
       
For Each WS In Worksheets

   Cells(1, 10) = "ticker"
   Cells(1, 11) = "yearly change"
   Cells(1, 12) = "pct change"
   Cells(1, 13) = "total volume"
     
   
    For I = 2 To rowcount + 1
    
        If WS.Cells(I, 1).Value <> WS.Cells(I - 1, 1).Value Then
        
            ticker = WS.Cells(I, 1).Value
            
            volume = WS.Cells(I, 7).Value
            
            open_price = WS.Cells(I, 3).Value
                     
            Debug.Print (ticker)
            Debug.Print (I)
        
        ElseIf WS.Cells(I + 1, 1).Value <> WS.Cells(I, 1).Value Then
                
            close_price = WS.Cells(I, 6).Value
                       
            volume = volume + WS.Cells(I, 7).Value
            
            yearly_change = close_price - open_price
                                      
            ticker = WS.Cells(I - 1, 1).Value
            
            WS.Range("J" & summary_row).Value = ticker
            
            WS.Range("m" & summary_row).Value = volume
            
            If yearly_change = 0 And open_price = 0 Then
                                        
                 WS.Range("k" & summary_row).Value = 0
            
                 WS.Range("l" & summary_row).Value = 0
                                                                                                                              
                 summary_row = summary_row + 1
                            
             ElseIf open_price <> 0 Then
                           
                
                WS.Range("k" & summary_row).Value = yearly_change
            
                WS.Range("l" & summary_row).Value = (yearly_change / open_price)
                                                                                                              
                summary_row = summary_row + 1
                            
            End If
                        
        ElseIf WS.Cells(I + 1, 1).Value = WS.Cells(I, 1).Value Then
            
            volume = volume + WS.Cells(I, 7).Value
                
        End If
        
    
        If WS.Cells(I, 11).Value > 0 Then
       
             WS.Cells(I, 11).Interior.ColorIndex = 4
        
        ElseIf WS.Cells(I, 11).Value < 0 Then
       
            WS.Cells(I, 11).Interior.ColorIndex = 3
   
    End If
    
             
       Next I
       
  WS.Range("L:L").NumberFormat = "0.00%"
    
       
    Next WS
    
End Sub


