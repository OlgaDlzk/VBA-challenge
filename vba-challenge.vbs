Sub stock_counter()

For Each ws In Worksheets

    ' Set up initial variables

    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Value As String
    Dim Open_Value As Double
    Dim Close_Value As Double
    
    Total_Value = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    
    
    ' Variables for bonus tasks
    
    Dim Gr_Inc_Name As String
    Dim Gr_Dec_Name As String
    Dim Gr_Tot_Vol_Name As String
    Dim Gr_Inc_Value As Double
    Dim Gr_Dec_Value As Double
    Dim Gr_Tot_Value As String
      
    Gr_Inc_Name = ""
    Gr_Dec_Name = ""
    Gr_Tot_Vol_Name = ""
    Gr_Inc_Value = 0
    Gr_Dec_Value = 0
    Gr_Tot_Value = 0

    ' Loop through all the stocks

    For i = 2 To lastrow
    
        ' Apply the condition
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Total_Value = Total_Value + ws.Cells(i, 7).Value
            Close_Value = ws.Cells(i, 6).Value
            Yearly_Change = Close_Value - Open_Value
            Percent_Change = Yearly_Change / Open_Value
            
            ' Finding the greatest percentange increase, decrease and the greatest total volume
            
            If Percent_Change > Gr_Inc_Value Then
            
                Gr_Inc_Value = Percent_Change
                Gr_Inc_Name = Ticker
                
            End If
                
            If Percent_Change < Gr_Dec_Value Then
            
                Gr_Dec_Value = Percent_Change
                Gr_Dec_Name = Ticker
                
            End If
                
            If CDec(Total_Value) > CDec(Gr_Tot_Value) Then
            
                Gr_Tot_Value = Total_Value
                Gr_Tot_Vol_Name = Ticker
                
            End If
            
                      
            ' Print the results in corresponding columns
            
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("L" & Summary_Table_Row).Value = Total_Value
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            ws.Range("K" & Summary_Table_Row).Value = Format(Percent_Change, "0.00%")
                                                                         
                   
            ' Reset the total value
            
            Total_Value = 0
                        
             ' Add one to the summary table row
            
            Summary_Table_Row = Summary_Table_Row + 1
            
        
        ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Open_Value = ws.Cells(i, 3).Value
            Total_Value = Total_Value + ws.Cells(i, 7).Value
       
       
        Else
       
            Total_Value = Total_Value + ws.Cells(i, 7).Value
            
            
       End If
        
    Next i
    
    ' printing out the greatest percentange increase, decrease and the greatest total volume
    
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 16).Value = Gr_Inc_Name
    ws.Cells(2, 17).Value = Format(Gr_Inc_Value, "0.00%")
    ws.Cells(3, 16).Value = Gr_Dec_Name
    ws.Cells(3, 17).Value = Format(Gr_Dec_Value, "0.00%")
    ws.Cells(4, 16).Value = Gr_Tot_Vol_Name
    ws.Cells(4, 17).Value = Gr_Tot_Value
    
        
    ' conditional formatting loop
    
    For i = 2 To lastrow
    
    ' Apply the condition
        
        If ws.Cells(i, 10).Value > 0 And ws.Cells(i, 11).Value > 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        Else
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            ws.Cells(i, 11).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    
    

Next ws

        MsgBox ("Operation complete")

End Sub
