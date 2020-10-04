Attribute VB_Name = "Module1"
Option Explicit

Sub VBA_stocks()

    Dim Ws As Worksheet
    
    For Each Ws In Worksheets
    
        Dim Ticker_Name, Current, Future, Max_volume_ticker, Max_Ticker_name, Min_Ticker_name As String
        Dim Total_Ticker_Volume, Open_Price, Close_Price, Yearly_Change, Percentage, Max_Percentage, Min_Percentage, Max_Volume As Double
        Dim row, Lastrow, i As Long
    
        ' Assign variables
        Ticker_Name = " "
        Open_Price = 0
        Close_Price = 0
        Yearly_Change = 0
        Percentage = 0
        Total_Ticker_Volume = 0
        Max_Ticker_name = " "
        Min_Ticker_name = " "
        Max_Percentage = 0
        Min_Percentage = 0
        Max_volume_ticker = " "
        Max_Volume = 0
    
        ' count rows,
        row = 2
        Lastrow = Ws.Cells(Rows.count, 1).End(xlUp).row
        

            ' creating the headers
        Ws.Range("I1").Value = "Ticker"
        Ws.Range("J1").Value = "Yearly Change"
        Ws.Range("K1").Value = "Percent Change"
        Ws.Range("L1").Value = "Total Stock Volume"
        Ws.Range("O2").Value = "Greatest % Increase"
        Ws.Range("O3").Value = "Greatest % Decrease"
        Ws.Range("O4").Value = "Greatest Total Volume"
        Ws.Range("P1").Value = "Ticker"
        Ws.Range("Q1").Value = "Value"
        
        'assign the first open price
        Open_Price = Ws.Cells(2, 3).Value
            
        For i = 2 To Lastrow ' loop through each rows
            Current = Ws.Cells(i, 1).Value
            Future = Ws.Cells(i + 1, 1).Value
            
            If Current <> Future Then
            
                Ticker_Name = Current
                Close_Price = Ws.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
                
                If Open_Price <> 0 Then
                    Percentage = (Yearly_Change / Open_Price) * 100
                Else
                    MsgBox ("The Open Price is 0 , please fix it.")
                End If

                Ws.Range("I" & row).Value = Ticker_Name
                Ws.Range("J" & row).Value = Yearly_Change
          
                If (Yearly_Change > 0) Then
                    Ws.Range("J" & row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    Ws.Range("J" & row).Interior.ColorIndex = 3
                End If

                Ws.Range("K" & row).Value = (CStr(Percentage) & "%")
                Ws.Range("L" & row).Value = Total_Ticker_Volume
                
                row = row + 1
                Yearly_Change = 0
                Close_Price = 0
                Open_Price = Ws.Cells(i + 1, 3).Value
              
                
                If (Percentage > Max_Percentage) Then
                    Max_Percentage = Percentage
                    Max_Ticker_name = Ticker_Name
                ElseIf (Percentage < Min_Percentage) Then
                    Min_Percentage = Percentage
                    Min_Ticker_name = Ticker_Name
                End If
                       
                If (Total_Ticker_Volume > Max_Volume) Then
                    Max_Volume = Total_Ticker_Volume
                    Max_volume_ticker = Ticker_Name
                End If
                Percentage = 0
                Total_Ticker_Volume = 0
                
            Else
    
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
            End If
      
        Next i
            
                Ws.Range("Q2").Value = (CStr(Max_Percentage) & "%")
                Ws.Range("Q3").Value = (CStr(Min_Percentage) & "%")
                Ws.Range("P2").Value = Max_Ticker_name
                Ws.Range("P3").Value = Min_Ticker_name
                Ws.Range("Q4").Value = Max_Volume
                Ws.Range("P4").Value = Max_volume_ticker
            
        
     Next Ws
End Sub

