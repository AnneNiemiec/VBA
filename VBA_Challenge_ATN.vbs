Sub Stock_Data()
    
    ' 1. Declare and Set Variables
    
    Dim ticker As String
    Dim ticker_date As Long
    Dim open_price As Double
    Dim high As Double
    Dim low As Double
    Dim close_price As Double
    Dim vol As Double
    Dim column As Integer
    Dim row As Integer
    Dim ticker_symbol As String
    Dim Annual_change As Double
    Dim percent_change As Double
    Dim Calculations As Long
    Dim i As Long
    Dim Total_stock_volume As Double
    
    Total_stock_volume = 0
    open_price = Cells(2, 3).Value
    close_price = 0
    Calculations = 2
    
lastrow = Cells(Rows.Count, 1).End(xlUp).row

'Loop through all stocks
For i = 2 To lastrow
    Total_stock_volume = Total_stock_volume + Cells(i, 7).Value
 
'Check if the stock ticker name is same

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
  
'Populate Ticker, Yearly Change, Percent Change and Total Stock Volumn Headers
    Range("I1").Value = "Ticker_Symbol"
    Range("J1").Value = "Annual_Change"
    Range("K1").Value = "Percent_Change"
    Range("L1").Value = "Total_stock_volume"
    Range("M1").Value = "open_price"
    Range("N1").Value = "close_price"
    Range("I1:N1").Font.Bold = True
           
    
'Set variables for each ticker
'Set ticker name
    ticker_symbol = Cells(i, 1).Value
    
'TODO: Set close_price
close_price = Cells(i, 6).Value

'TODO: Set Annual_change
' Hint: close_price - open_price
Annual_change = close_price - open_price

'TODO: Set percent_change
' Hint: (close_price - open_price) / open_price
If open_price = 0 Then
    percent_change = 0
    Else
    percent_change = (close_price - open_price) / open_price

    End If
    
'Update Calculations
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Print the stock ticker in the calculations table
    Range("I" & Calculations).Value = ticker_symbol
    
'Print the open_price to calculations
    Range("M" & Calculations).Value = open_price
    
'Print the close_price to calculations
    Range("N" & Calculations).Value = close_price
    
'Print Annual_change to calculations
    Range("J" & Calculations).Value = Annual_change
    Annual_change = (close_price - open_price)
    
'conditional formatting
    If Range("J" & Calculations) > 0 Then
     Range("J" & Calculations).Interior.ColorIndex = 4
        ElseIf Range("J" & Calculations) < 0 Then
         Range("J" & Calculations).Interior.ColorIndex = 3
            Else
           Range("J" & Calculations).Interior.ColorIndex = 0
           End If
                      
        
'Print Percent_change to calculations
    Range("K" & Calculations).Value = percent_change
    Range("K" & Calculations).NumberFormat = "0.00%"
    
    
'Print total_stock_volume to calculations
    Range("L" & Calculations).Value = Total_stock_volume
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TODO: Set open_price for next ticket
' Hint: Cells(i+1, 3).Value
open_price = Cells(i + 1, 3).Value

'Add one to the calculations row
    Calculations = Calculations + 1
    
'Reset the ticker total
    Total_stock_volume = 0

    
'If the cell immediately following a row is the same...
    Else
    If open_price = 0 Then
    open_price = Cells(i, 3).Value
    End If
    
    
         
     

            
      
        End If
    
    Next i
   
End Sub

