Attribute VB_Name = "Module1"
Sub VBAchallenge()
    'Declaring the varables
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Vol As Double
    Dim tablerow As Double
    Dim Greatest_Per As Double
    Dim Smallest_Per As Double
    Dim Greatest_Total As Double
    Dim Lastrow As Double
    Dim Lastcol As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    
    'For loop to work on each sheet
    For Each ws In Worksheets
       Worksheets(ws.Name).Activate
        
        'determine the Last Row and the last column
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        Lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
       
        
        'Set up the Column Headings
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("K2:K1000000").NumberFormat = "0.00%"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
         
         'Assigning Variables
        'Greatest_Per = 0
        'Smallest_Per = 0
        'Greatest_Total = 0
         tablerow = 2
               
        'determining the opening Price
        Open_Price = Cells(2, 3).Value
        Total_Vol = 0
        
        
        
        'run For loop to iterate through all the rows
        For i = 2 To Lastrow
        
            'check if the tickers are different
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
                'find the ticker value and assign it
                Ticker = Cells(i, 1).Value
                Cells(tablerow, Lastcol + 2).Value = Ticker
                
                'calculate and polulate the yearly price change
                Close_Price = Cells(i, Lastcol - 1).Value
                Yearly_Change = Close_Price - Open_Price
                Cells(tablerow, Lastcol + 3).Value = Yearly_Change
                
                'Apply the color formatting
                
                If Cells(tablerow, Lastcol + 3).Value <= 0 Then
                   Cells(tablerow, Lastcol + 3).Interior.ColorIndex = 3
                    
                Else
                   Cells(tablerow, Lastcol + 3).Interior.ColorIndex = 4
                
                End If
                
                'Calculate the Percentage of Yearly Price change
                If Open_Price <> 0 Then
                    Percent_Change = Yearly_Change / Open_Price
                Else
                    Percent_Change = Yearly_Change
 
                End If
                    
                Cells(tablerow, Lastcol + 4).Value = Percent_Change
                
                'calculate and assign the Total Volume
                Total_Vol = Total_Vol + Cells(i, Lastcol).Value
                Cells(tablerow, Lastcol + 5).Value = Total_Vol


                ' reset data before move into next table
                tablerow = tablerow + 1
                Total_Vol = 0
                Open_Price = Cells(i + 1, 3).Value

               'Add the row volumes if the ticker is same
            Else
                Total_Vol = Total_Vol + Cells(i, 7).Value

            End If
            
            
        Next i
        
        

    Next ws

End Sub
