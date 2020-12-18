Sub Stock_Data()


' Loop through all sheets of Worksheet
For Each ws In Worksheets
    
    
    ' Create a header
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
    
    ' Autofit to display data for Ticker,Yearly Change,Percent Change,Total Stock Volume
        ws.Columns("I:L").AutoFit
        
    
    ' Declare the required variables
        Dim Counter As Integer
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Ticker As String
        Dim Yearly_change As Double
        Dim Percent_change As Double
        Dim Tot_Stock As Double
        
           
    ' Get the last row from sheet
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' Initialize the variables
        Counter = 0
        j = 2
        Tot_Stock = 0
      
    
          ' Loop through all Ticker's
            For i = 2 To lastrow
            
              
                ' Stop when we have reached end of the row for each Ticker and collect Closing price
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
                    ' Dump the Closing price for each Ticker into Close_Price variable
                    Close_Price = ws.Cells(i, 6).Value
                    
                    
                    ' Calculate Yearly change
                    Yearly_change = Close_Price - Open_Price
                    ws.Cells(j, 10).Value = Yearly_change ' Yearly Change
                    
                    
                    ' Calculate Percent change
                    Percent_change = Yearly_change / IIf(Open_Price = 0, 1, Open_Price)
                    ws.Cells(j, 11).Value = Format(Percent_change, "Percent") ' Percent Change
                    
                    
                    ' Calculate Total stock volume for each Ticker when it reaches last row
                    Tot_Stock = Tot_Stock + ws.Cells(i, 7).Value
                    ws.Cells(j, 12).Value = Tot_Stock ' Total stock
                    
                    
                    ' Increment j
                    j = j + 1
                    
                    
                    ' Reset the Counter
                    Counter = 0
                    
                    
                    ' Reset the Tot_Stock
                    Tot_Stock = 0
            
            
                Else
                
                    ' All the below code should every time when we have same Ticker during comparison of current row and next row
            
                    ' Increment the Row Counter
                    Counter = Counter + 1
                    
                    ' Fetch the Opening price from 1st row for each Ticker and populate Ticker values
                    If Counter = 1 Then
                    Open_Price = ws.Cells(i, 3).Value ' Open_Price
                    ws.Cells(j, 9).Value = ws.Cells(i, 1) ' Ticker
                    End If
                    
                    ' Calculate Total stock volume for each Ticker
                    Tot_Stock = Tot_Stock + ws.Cells(i, 7).Value
                    
             
                End If
            
            
            Next i
      
      
          ' Conditional Formatting to highlight Positive and Negative numbers for Yearly Change
            For r = 2 To lastrow
                
                ' Note: The IsEmpty will make sure we are not printing highlighting blank cells
                If IsEmpty(ws.Cells(r, 10).Value) = False Then
                    If (ws.Cells(r, 10).Value < 0) Then
                        ws.Cells(r, 10).Interior.ColorIndex = 3 ' Red
                    Else
                        ws.Cells(r, 10).Interior.ColorIndex = 4 ' Green
                    End If
                End If
            
            Next r
            
            
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ' Bonus Part
    ' Create a header
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
        
        
    ' Declare the required variables
    Dim Max_P_Inc As Double
    Dim Min_P_Dec As Double
    Dim Max_P_Vol As Double
    Dim Max_P_Inc_Tkr As String
    Dim Max_P_Dec_Tkr As String
    Dim Max_P_Vol_Tkr As String
    
    
       
    ' Find out the Greatest % increase
    ' Initialize the Max Variable
    Max_P_Inc = 0
        
        
        For r = 2 To lastrow
            
            If IsEmpty(ws.Cells(r, 11).Value) = False Then
                If (ws.Cells(r, 11).Value > Max_P_Inc) Then
                    Max_P_Inc = ws.Cells(r, 11)
                    Max_P_Inc_Tkr = ws.Cells(r, 9)
                End If
            End If
        
        Next r
            
        ws.Range("P2").Value = Max_P_Inc_Tkr ' Ticker
        ws.Range("Q2").Value = Format(Max_P_Inc, "Percent") ' Greatest % increase
        
        
        
    ' Find out the Greatest % decrease
    ' Initialize the Max Variable
    Min_P_Dec = 0
        
        
        For r = 2 To lastrow
            
            If IsEmpty(ws.Cells(r, 11).Value) = False Then
                If (ws.Cells(r, 11).Value < Min_P_Dec) Then
                    Min_P_Dec = ws.Cells(r, 11)
                    Max_P_Dec_Tkr = ws.Cells(r, 9)
                End If
            End If
        
        Next r
            
        ws.Range("P3").Value = Max_P_Dec_Tkr ' Ticker
        ws.Range("Q3").Value = Format(Min_P_Dec, "Percent") ' Greatest % decrease
        
        
        
    ' Find out the Greatest total volume
    ' Initialize the Max Variable
    Max_P_Vol = 0
        
        
        For r = 2 To lastrow
            
            If IsEmpty(ws.Cells(r, 12).Value) = False Then
                If (ws.Cells(r, 12).Value > Max_P_Vol) Then
                    Max_P_Vol = ws.Cells(r, 12)
                    Max_P_Vol_Tkr = ws.Cells(r, 9)
                End If
            End If
        
        Next r
            
        ws.Range("P4").Value = Max_P_Vol_Tkr ' Ticker
        ws.Range("Q4").Value = Max_P_Vol ' Greatest total volume
        ws.Range("Q4").NumberFormat = "0.0000E+00" ' Greatest total volume
    
    
    ' Autofit to display data
    ws.Columns("O:Q").AutoFit
    

' End of the Loop for all sheets of Multiple_year_stock_data
Next ws
  
End Sub

