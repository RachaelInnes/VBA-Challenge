Attribute VB_Name = "Module1"
Sub CalculateTickerData()

    ' Declare variables
    Dim Ticker_Name As String
    Dim Total_Stock_Volume As Double
    Dim Quarterly_Change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim price_change_percent As Double
    Dim lastRow As Long
    Dim Summary_Table_Row As Long
    Dim i As Long
    Dim wb As Workbook
    Dim ws As Worksheet
       
    
    ' Reference to the current workbook
    Set wb = ThisWorkbook

    ' Initialize variables
    Total_Stock_Volume = 0
    Summary_Table_Row = 2
    
    
    Set wb = ThisWorkbook ' Reference to the current workbook
    
    ' Loop through each worksheet
    For Each ws In wb.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
      
    
     ' Initialize open_price for the first ticker
     open_price = Cells(2, 3).Value
          
    ' Loop through all Ticker volumes
    For i = 2 To lastRow
                
        ' Check if the volume cell is numeric
        If IsNumeric(Cells(i, 7).Value) Then
            'Add to the Total Stock Volume
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                        
            ' Check if we are still within the same Ticker or at the last row
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Or i = lastRow Then
            ' Set the ticker name
            Ticker_Name = Cells(i, 1).Value
                                             
                ' Set the close price
                close_price = Cells(i, 6).Value
                
                ' Calculate Quarterly change in Price
                price_change = close_price - open_price
                If open_price <> 0 Then
                    price_change_percent = (price_change / open_price) / 100 * 100
                                                 
                Else
                    price_change_percent = 0
                End If
                           
                           
                ' Print the Ticker Brand in the Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' Print the Quarterly Change Amount to the Summary Table
                Range("J" & Summary_Table_Row).Value = price_change
                
                 ' Print the Percent Change Percent to the Summary Table
                Range("K" & Summary_Table_Row).Value = price_change_percent

                ' Print the Ticker Total Stock Volume Amount to the Summary Table
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                            
                               
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset the Total Stock Volume
                Total_Stock_Volume = 0
                
                ' Reset the open price for the next ticker if not the last row
                If i + 1 <= lastRow Then
                    open_price = Cells(i + 1, 3).Value
                End If
            
            End If
                              
        Else
         ' Handle non-numeric volume value
            Debug.Print "Non-numeric volume at Row: " & i & ", Value: " & Cells(i, 7).Value
        End If
    Next i
 
     ' Clear any existing conditional formatting
    ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Delete
    
    ' Add new conditional formatting rule
    With ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
        .Interior.Color = RGB(0, 255, 0)
        
    End With
    
             
    ' Add new conditional formatting rule for negative percentage changes (red)
    With ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 0, 0)
    
    End With
    
        ' Format the Percent Change column (K) to show percentages with two decimal places
    ws.Range("K2:K" & Summary_Table_Row - 1).NumberFormat = "0.00%"
    
   Next ws
   
End Sub

    Sub FindMaxVolumeForAllSheets():
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxVolume As Double
    Dim maxTicker As String
    
    ' Reference to the current workbook
    Set wb = ThisWorkbook
     
    'Loop through each worksheet
    For Each ws In wb.Worksheets
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
      
    ' Find the highest volume and corresponding ticker symbol
    maxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
    maxTicker = ws.Cells(Application.WorksheetFunction.Match(maxVolume, ws.Range("L2:L" & lastRow), 0) + 1, "I").Value
    
    
    ' Write the results to the summary table
    ws.Range("P4").Value = maxTicker
    ws.Range("Q4").Value = maxVolume
    
       
   Next ws
   
End Sub

 Sub Greatestpercentchange()
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim Greatestpercentchange As Double
    Dim maxTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column K
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        
        ' Ensure there is data to process
        If lastRow > 1 Then
            ' Find the greatest percent change and the corresponding ticker symbol
            Greatestpercentchange = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
            maxTicker = ws.Cells(Application.WorksheetFunction.Match(Greatestpercentchange, ws.Range("K2:K" & lastRow), 0) + 1, "I").Value
            
            ' Write the results to the summary table
            ws.Range("P2").Value = maxTicker
            ws.Range("Q2").Value = Greatestpercentchange
        
        End If
    Next ws
    
End Sub

Sub smallestpercentchange():

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim smallestpercentchange As Double
    Dim minTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in column K
        lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        
        ' Ensure there is data to process
        If lastRow > 1 Then
            ' Find the greatest percent change and the corresponding ticker symbol
            smallestpercentchange = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
            minTicker = ws.Cells(Application.WorksheetFunction.Match(smallestpercentchange, ws.Range("K2:K" & lastRow), 0) + 1, "I").Value
            
            ' Write the results to the summary table
            ws.Range("P3").Value = minTicker
            ws.Range("Q3").Value = smallestpercentchange
        
         ' Format cells as percentage
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
        
        End If
    Next ws
    
End Sub
       
    



    
    
  
   


    
