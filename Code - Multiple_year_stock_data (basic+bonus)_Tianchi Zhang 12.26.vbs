Attribute VB_Name = "StockAnalysis_Final"
Sub Run_1_StockAnalysis()


'Set an initial variable for holding the ticker symbol of each company
Dim Ticker As String

'Set an initial variable for holding the total stock volume per ticker symbol
    'P.S. I'm using Variant for volume since the "Long" variable didn't work...
Dim Ticker_Total_Volume As Variant
Ticker_Total_Volume = 0

'Set initial variables for holding the opening and closing prices
    'which are on the first day of a year and the last day of a year respectively for each particular ticker symbol
    'P.S. I'm using Variant for all price related variables since my Mac is having trouble running with "Double" variables...
Dim Open_Price As Variant
Dim Close_Price As Variant
    'Setting the first opening price to that of the very first ticker symbol
Open_Price = Range("C2").Value

'Set an initial variable for holding the yearly change of a ticker symbol
Dim Yearly_Change As Variant
'Set an initial varialbe for holding the percent change of a ticker symbol
Dim Percent_Change As Variant

'Keep track of the location for each ticker symbol in the summary table
Dim Summary_Row As Integer
Summary_Row = 2

'Set an variable for holding the last row number of the raw data in a worksheet
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all stock prices for a particular year
For i = 2 To lastrow

    'Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Set the ticker symbol
        Ticker = Cells(i, 1).Value
        
        'Add to the ticker total volume
        Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
        
        'Print the ticker symbol in the summary table
        Range("I" & Summary_Row).Value = Ticker
        
        'Print the total volume of the ticker being summarized to the summary table
        Range("L" & Summary_Row).Value = Ticker_Total_Volume
        
        'Set the end price of a given year for a particular ticker symbol for calculations of change
        Close_Price = Cells(i, 6).Value
        
        'Calculate the yearly change and the percent change
        Yearly_Change = Close_Price - Open_Price
        'P.S. e.g. Sticker symbol "PLNT" has all pricing information as 0, thus a conditional needs to be customized...
            'Check if a ticker has a price at the beginning and end of the year, if not...
            If Close_Price = 0 And Open_Price = 0 Then
                'Make percent change 0 as there's no data for this hypothetical ticker
                Percent_Change = 0
            'Check if a ticker entered the market in the middle of a year which typically results in it not having a open price
            ElseIf Close_Price <> 0 And Open_Price = 0 Then
                Percent_Change = 0
                Yearly_Change = "Override"
                'Use a message box to call out this type of ticker everytime Excel finds one so the user can write it down for edits later...
                MsgBox ("Attention! Ticker " & Cells(i, 1) & " just entered the market this year. Calculate the Yearly Change and the Percent Change manually before moving on to complete Button 1 or before clicking on Button 2 to generate the brief summary report.")
            'If everything is "normal" or in a typical situation, simply calculate the percent change...
            Else
                Percent_Change = (Close_Price - Open_Price) / Open_Price
            End If
        
        'Print the yearly change and percent change of the ticker in the summary table
        Range("J" & Summary_Row).Value = Yearly_Change
        Range("K" & Summary_Row).Value = FormatPercent(Str(Percent_Change), 2)
        
        'Add conditional formating to highlight positive change in green and negative change in red
        If Yearly_Change > 0 Then
            Range("J" & Summary_Row).Interior.Color = RGB(146, 208, 80)
        Else
            Range("J" & Summary_Row).Interior.Color = RGB(248, 111, 108)
        End If
        
        'Set the new opening price of a given year for a particular ticker symbol for calculations of change
        Open_Price = Cells(i + 1, 3).Value
    
        'Move down one row in the summary table to summarize the next ticker
        Summary_Row = Summary_Row + 1
        
        'Reset the ticker total volume so we can start summarizing again for the next ticker symbol
        Ticker_Total_Volume = 0
        
    'If the cell immediately following a row is the same ticker symbol, we need to keep adding the volume to the total
    Else
    
        'Add to the ticker total volume
        Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
        
    End If

Next i
        

End Sub


Sub Run_2_StockSummary_Bonus()

'Set an variable for holding the last row number of the first summary table in a worksheet
Dim lastrow2 As Long
lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row

'Set an initial variable for holding the array of yearly changes
Dim Percent_Change_Array() As Variant
    'Set a dynamic array with the number of items in the percent change array so that module can be run on different sheets in case the numbers of ticker symbols are different
    ReDim Percent_Change_Array(2 To lastrow2)

'Set an initial variable for holding the array of total stock volumes
Dim Total_Stock_Volume_Array() As Variant
    'Set a dynamic array with the number of items in the total stock array so that module can be run on different sheets in case the numbers of ticker symbols are different
    ReDim Total_Stock_Volume_Array(2 To lastrow2)

'Set variables for holding the maximum summary values for each category
Dim MaxPercent As Variant
Dim MinPercent As Variant
Dim MaxVolume As Variant

'Loop through all items under Percent Change and Total Stock Volume, and add them to the defined arrays
For i = 2 To lastrow2
    
    'Add items to the percent change array
    Percent_Change_Array(i) = Cells(i, 11).Value
    'Add items to the total stock volume array
    Total_Stock_Volume_Array(i) = Cells(i, 12).Value
    
Next i

'Set required statistical summaries to defined maximum and minimum variables
MaxPercent = WorksheetFunction.Max(Percent_Change_Array)
MinPercent = WorksheetFunction.Min(Percent_Change_Array)
MaxVolume = WorksheetFunction.Max(Total_Stock_Volume_Array)

'Loop through all Ticker&Percent concatenation to find the matches with the values in the summary table so we can return the ticker symbol
For i = 2 To lastrow2
    
    'Check if a specific row has the matching information we are looking for; in this case, for Greatest % Increase...
    If Cells(i, 9) & Cells(i, 11) = Cells(i, 9) & MaxPercent Then

        'Print the value and its matching ticker symbol in the summmary table
        Range("P2").Value = Cells(i, 9).Value
            'Print the value in the form of percentage
            Range("Q2").Value = FormatPercent(Str(MaxPercent), 2)
        
    'For Greatest % Decrease...
    ElseIf Cells(i, 9) & Cells(i, 11) = Cells(i, 9) & MinPercent Then
    
        'Print the value and its matching ticker symbol in the summmary table
        Range("P3").Value = Cells(i, 9).Value
            'Print the value in the form of percentage
            Range("Q3").Value = FormatPercent(Str(MinPercent), 2)
        
    'For Greatest Total Volume...
    ElseIf Cells(i, 9) & Cells(i, 12) = Cells(i, 9) & MaxVolume Then
        
        'Print the value and its matching ticker symbol in the summmary table
        Range("P4").Value = Cells(i, 9).Value
        Range("Q4").Value = MaxVolume
        
    End If
    
Next i


End Sub


