Attribute VB_Name = "Module1"
Sub Mod2Challenge():
    
    ' set place holder to set the loop for each worksheet
    
    Dim ws As Worksheet

    For Each ws In Worksheets
    
    ws.Activate
    ' Defining variable to isolate ticker symbols
    
        Dim Ticker_Symbol As String
        
        ' Setting a variable to hold the total stock volume
        
        Dim Stock_Total As Double
        
        ' Setting Variable to hold date
        
        Dim Yearly_Change As Double
        
        ' set place holder for stock table
        
        Dim Stock_Table As Double
            
            Stock_Table = 2
        
        ' set place holder for close amount for the year
        Dim Close_Amount As Double
        
            Close_Amount = 2
        
        ' set place holder for open amount for the year
        Dim Open_Amount As Double
        
            Open_Amount = 2
        
        ' set place holder for amount changed from 01/02 to 12/31
        Dim Year_Change As Double
            
            Year_Change = 2
            
        ' set place holder for percent change for the ticker
        Dim Percent_Change As Double
        
            Percent_Change = 2
        
        ' set place holder to set the loop for each worksheet

        
        ' set place holder for Greatest Percent Increase
        Dim Greatest_Increase As Double
        
        ' set data type for Greatest Percent Increase
        Dim Increase_Cell As Double
        
            Increase_Cell = 2
        
        ' set data type for Greatest Percent Decrease
        Dim Greatest_Decrease As Double
        
        ' set cell to hold greatest percent decrease
        Dim Decrease_Cell As Double
        
            Decrease_Cell = 3
            
        ' set data type for Greatest Volume
        Dim Greatest_Volume As Double
        
        ' set place holder for Greatest Volume
        Dim Greatest_Cell As Double
        
            Greatest_Cell = 4
        
        Dim Ticker_Greatest_Volume As String
        
        Dim Ticker_Greatest_Increase As String
        
        Dim Ticker_Greatest_Decrease As String
            
        
        ' Set column names for data set
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"
        Range("N1").Value = "Open Amount"
        Range("O1").Value = "Close Amount"
        Range("R1").Value = "Ticker"
        Range("S1").Value = "Value"
        Range("L:L").NumberFormat = "0.00%"
        Range("S2:S3").NumberFormat = "0.00%"
        Range("S4").NumberFormat = "0"
        Range("K:K").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Range("Q2").Value = "Greatest % Increase"
        Range("Q3").Value = "Greatest % Decrease"
        Range("Q4").Value = "Greatest Total Volume"
    
        ' Create step to reference worksheet name
         ' Set A = Worksheets("2018")
        
        
        ' Identify the last row
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
                
                ' Set condition for yearly change and percent change greater than zero
        For x = 2 To lastRow
    
            If ws.Cells(x, "L").Value > 0 Then
            
                ws.Cells(x, "L").Interior.Color = rgbForestGreen
                ws.Cells(x, "K").Interior.Color = rgbForestGreen
        
            ElseIf ws.Cells(x, "L").Value < 0 Then
                
                ws.Cells(x, "L").Interior.Color = rgbIndianRed
                ws.Cells(x, "K").Interior.Color = rgbIndianRed
                
            Else
                ws.Cells(x, "L").Interior.Color = rgbGreenYellow
                ws.Cells(x, "K").Interior.Color = rgbGreenYellow
            
            End If
        
        Next x
              
                ' Creating loop to capture all unique ticker symbols and sum all stock volume to their respective ticker
     
        For x = 2 To lastRow
                
            ' Review Ticker_Symbol
                    
            If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
                    
                ' Set Ticker Symbol
                                
                Ticker_Symbol = Cells(x, 1).Value
                                
                ' Total Stock Volume
                                
                Stock_Total = Stock_Total + Cells(x, 7).Value
                            
                    ' Transfer Ticker Symbol to Column J
                            
                Range("J" & Stock_Table).Value = Ticker_Symbol
                                
                ' Transfer Stock Volume to Column M
                                
                Range("M" & Stock_Table).Value = Stock_Total
                                
                                
                ' Append to Stock Table
                                
                Stock_Table = Stock_Table + 1
                                
                ' Continue with Stock Total
                                
                Stock_Total = 0
                                
                ' Add matching Tick_Symbol Stock Volumes
                                 
            Else
                    
                Stock_Total = Stock_Total + Cells(x, 7).Value
                        
            End If
                
        Next x
                
                ' Isolate the close amount for the year, ending 12/31
        For x = 2 To lastRow
        
            If Cells(x + 1, 1).Value <> Cells(x, 1).Value And Cells(x, 2).Value = 20201231 Or Cells(x + 1, 1).Value <> Cells(x, 1).Value And Cells(x, 2).Value = 20191231 Or Cells(x + 1, 1).Value <> Cells(x, 1).Value And Cells(x, 2).Value = 20181231 Then
                
            
                ' Set Ticker Symbol
                
                Ticker_Symbol = Cells(x, 1).Value
                
                Close_Value = Cells(x, 6).Value
                
                Range("O" & Close_Amount).Value = Close_Value
                
                Close_Amount = Close_Amount + 1
                        
            End If
                        
        Next x
        
        ' Isolate the open amount for the year, beginning 01/02
        For x = 2 To lastRow
        
            If Cells(x + 1, 1).Value = Cells(x, 1).Value And Cells(x, 2).Value = 20200102 Or Cells(x + 1, 1).Value = Cells(x, 1).Value And Cells(x, 2).Value = 20190102 Or Cells(x + 1, 1).Value = Cells(x, 1).Value And Cells(x, 2).Value = 20180102 Then
            
                Ticker_Symbol = Cells(x, 1).Value
            
                Open_Value = Cells(x, 3).Value
            
                Range("N" & Open_Amount).Value = Open_Value
            
                Open_Amount = Open_Amount + 1
            
           End If
           
        ' Subtract close and open values to attain the yearly change
            
        Next x
        
        For x = 2 To lastRow
        
            If Cells(x, "N") <> Cells(x, "O") Or Cells(x, "N") = Cells(x, "O") And Cells(x, "N").Value <> 0 Then
            
            
                Yearly_Change = Cells(x, "O").Value - Cells(x, "N").Value
                
                Range("K" & Year_Change).Value = Yearly_Change
                
                Percentage_Change = Cells(x, "K").Value / Cells(x, "N")
                
                Range("L" & Percent_Change).Value = Percentage_Change
                
                Year_Change = Year_Change + 1
                
                Percent_Change = Percent_Change + 1
                
                        
            End If
            
            
        Next x
        
        ' Divide close and open values by the close value to attain the percent change
        
        For x = 2 To lastRow
        
            If Cells(x + 1, "J").Value <> Cells(x, "J").Value Then
            
                ' Set Ticker Symbol
                
                Ticker_Symbol = Cells(x, "J").Value
                
                ' Greatest % Change
                
                Greatest_Increase = WorksheetFunction.Max(Range("L:L"))
                
                
                ' Transfer Greatest Increase to Column R
                
                Range("S" & Increase_Cell).Value = Greatest_Increase
                
                Greatest_Decrease = WorksheetFunction.Min(Range("L:L"))
                
                Range("S" & Decrease_Cell).Value = Greatest_Decrease
                
                Greatest_Volume = WorksheetFunction.Max(Range("M:M"))
                
                Range("S" & Greatest_Cell).Value = Greatest_Volume
                
    
            End If
            
        Next x
        
        Greatest_Ticker_Increase = Application.WorksheetFunction.XLookup(ws.Range("S2"), ws.Range("L2:L3050"), ws.Range("J2:J3050"), 0, 0, 1)
        ws.Range("R2").Value = Greatest_Ticker_Increase
       
        Greatest_Ticker_Decrease = Application.WorksheetFunction.XLookup(ws.Range("S3"), ws.Range("L2:L3050"), ws.Range("J2:J3050"), 0, 0, 1)
        ws.Range("R3").Value = Greatest_Ticker_Decrease
        
        Greatest_Ticker_Volume = Application.WorksheetFunction.XLookup(ws.Range("S4"), ws.Range("M2:M3050"), ws.Range("J2:J3050"), 0, 0, 1)
        ws.Range("R4").Value = Greatest_Ticker_Volume
        
        ws.Columns("A:Z").AutoFit
        
    Next ws
            
End Sub

