Attribute VB_Name = "Module1"
Sub Multiple_Year()

    'Create a script that will loop through all the stocks for one year and output the following info
        'Create the Ticker Symbol
        'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
        'Percentage of from opening price at the beginning of a given year to the closing price at the end of that year
        'The total stock volume of the stock
  
'Loop through all sheets
Dim ws As Worksheet
For Each ws In Worksheets

'Label each column for each worksheet
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"

  
'Set a variable for the ticker name
Dim TickerName As String

Dim Open_Price As Double
Open_Price = 0

Dim Close_Price As Double
Close_Price = 0

Dim Price_Change As Double
Price_Change = 0

Dim Percentage_Change As Double
Percentage_Change = 0

Dim Stock_Volume As Double
Stock_Volume = 0

'Challenge Homework

Dim Max_Ticker_Name As String
Max_Ticker_Name = ""

Dim Min_Ticker_Name As String
Min_Ticker_Name = ""

Dim Greatest_Percent_Increase As Double
Greatest_Percent_Increase = 0

Dim Greatest_Percent_Decrease As Double
Greatest_Percent_Decrease = 0

Dim Greatest_TotalVolume_Ticker As String
Greatest_TotalVolume_Ticker = ""

Dim Greatest_Total_Volume As Double
Greatest_Total_Volume = 0

'Setting Titles for Challenge
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

  
  
'Keep track of the location for each tickername in the summary table
Dim Summary_Table As Long

Summary_Table = 2

  
 'Setting to last row
 
Dim LastRow As Long
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  'Setting the value for Open Price for all worksheets
  Open_Price = ws.Cells(2, 3).Value
  
     
For i = 2 To LastRow

'Checking if we are within the same tickername

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Set the TickerName
        TickerName = ws.Cells(i, 1).Value
        
        'Calculating Yearly Change
        Close_Price = ws.Cells(i, 6).Value
        Price_Change = Close_Price - Open_Price
        

        
        'Calculate Percentage_Change
        If Open_Price <> 0 Then
            Percentage_Change = (Price_Change / Open_Price) * 100
            
    End If
        

        'Print the ticker name in the summary table
        ws.Range("I" & Summary_Table).Value = TickerName
        
        'Print Price_Change in the Summary Table under Yearly Change
        ws.Range("J" & Summary_Table).Value = Price_Change
        
            'Higlighting green for positive and red for negative changes
            If (Price_Change > 0) Then
                ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
            
            ElseIf (Price_Change <= 0) Then
                ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
            
            End If
                
        
        'Print Percentage_Change
        ws.Range("K" & Summary_Table).Value = (CStr(Percentage_Change) & "%")
        
        'Print Stock_Volume
        ws.Range("L" & Summary_Table).Value = Stock_Volume
        
        'Add one to the Summary_Table
        Summary_Table = Summary_Table + 1
        
        Price_Change = 0
        Close_Price = 0
        Open_Price = ws.Cells(i + 1, 3).Value
  
  'Challenge Calculations
  
  If (Percentage_Change > Greatest_Percent_Increase) Then
    Greatest_Percent_Increase = Percentage_Change
    Max_Ticker_Name = TickerName
    
ElseIf (Percentage_Change < Greatest_Percent_Decrease) Then
    Greatest_Percent_Decrease = Percentage_Change
    Min_Ticker_Name = TickerName
    
 End If
 

 
 If (Stock_Volume > Greatest_Total_Volume) Then
    Greatest_Total_Volume = Stock_Volume
    Greatest_TotalVolume_Ticker = TickerName
  End If
  
 Percentage_Change = 0
 Stock_Volume = 0
  
'Calculate Stock Volume
  
Else
Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
End If

 Next i
 
  ws.Range("P2").Value = Max_Ticker_Name
  ws.Range("P3").Value = Min_Ticker_Name
  ws.Range("P4").Value = Greatest_TotalVolume_Ticker
  ws.Range("Q2").Value = (CStr(Greatest_Percent_Increase) & "%")
  ws.Range("Q3").Value = (CStr(Greatest_Percent_Decrease) & "%")
  ws.Range("Q4").Value = Greatest_Total_Volume

 

Next ws


End Sub















