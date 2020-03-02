Sub VBA_HomeWork_Solution()

    'Looping through all worksheets
    For Each ws In Worksheets

        'Finding the Last Row of each sheet one at a time
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Setting all intial variables for a sheet
        
        Dim Ticker_Name As String
        Dim open_value As Double
        Dim close_value As Double
        Dim Change_Value As Double
        Dim Volume_Total As Double
        
        'Setting Volume_Total as 0 at beginning of each sheet
        Volume_Total = 0

        'Setting summary row variable to capture values for additional challanges
        Dim Ticker_Summary_Row As Integer
        
        'Setting Ticker_Summary_Row as 2 to start from second row
        Ticker_Summary_Row = 2
        
        'Hard coding summary row values as this are needed to set only one time per sheet
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 15).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        'Setting Open_value variable at beggining to get initial value before going through full sheet
        open_value = ws.Cells(2, 3).Value
        
        'ENtering For loop for each sheet
        For i = 2 To LastRow
 
        ' Checking if ticker name is changed to set variabes for previous ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                  'Set the closing value for previous ticker
                  
                  close_value = ws.Cells(i, 6).Value
            
                  'Add last volume value for previous tikcer before resetting it
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                  
                  'Capturing values for Annual ticker summary
                  Ticker_Name = ws.Cells(i, 1)
                  Change_Value = close_value - open_value
                  
                  'Setting Annual ticker summary in newly created rows
                  ws.Range("I" & Ticker_Summary_Row).Value = Ticker_Name
            
                  'Setting Change ticker value and filling with red/green if diff is negative or positive
                  If Change_Value < 0 Then
                  
                    ws.Range("J" & Ticker_Summary_Row).Value = Change_Value
                    ws.Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 3
                    
                  Else
                    
                    ws.Range("J" & Ticker_Summary_Row).Value = Change_Value
                    ws.Range("J" & Ticker_Summary_Row).Interior.ColorIndex = 4
                    
                End If
            
                   'Handling errors when opening value is 0 when calculating percent change
                   
                  If open_value <> 0 Then
                  
                    ws.Range("K" & Ticker_Summary_Row).Value = Change_Value / open_value
                    ws.Range("K" & Ticker_Summary_Row).NumberFormat = "0.00%"
                       
                    
                  Else
                  
                     ws.Range("K" & Ticker_Summary_Row).Value = 0
                    
                  End If
                  
                  'Storing totat in summary rows before resetting it for next ticker
                  ws.Range("L" & Ticker_Summary_Row).Value = Volume_Total
            
                  'Adding one to the summary table row for nect ticker
                  Ticker_Summary_Row = Ticker_Summary_Row + 1
                  
                  'Reset the volume total for next ticker
                  Volume_Total = 0
                  
                  'Setting opening Value for next ticker
                  open_value = ws.Cells(i + 1, 3).Value
            
                'If the ticker is not changed, keep adding volume to total volume
            Else

                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value

         End If

  Next i
  
        'Setting values for challenges after the row summary is completed but before looping over next worksheet
        
        'Defining additional variables to capture maximum, minimum change in percent and max volume
        
        Dim summary_row As Integer
        Dim per_range As String
        Dim vol_range As String
        Dim tkr_range As String
        
        summary_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        per_range = "K2:K" + Trim(Str(summary_row))
        vol_range = "L2:L" + Trim(Str(summary_row))
        tkr_range = "I2:I" + Trim(Str(summary_row))
        
        'Finding Maximum, Minimum percent change and Max Volume from summary data
        max_value = WorksheetFunction.Max(ws.Range(per_range))
        min_value = WorksheetFunction.Min(ws.Range(per_range))
        max_volume = WorksheetFunction.Max(ws.Range(vol_range))
        
        'Setting new values for challenges using Index and Matching function
        ws.Cells(2, 15) = WorksheetFunction.Index(ws.Range(tkr_range), WorksheetFunction.Match(max_value, ws.Range(per_range), 0))
        ws.Cells(2, 16) = max_value
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        ws.Cells(3, 15) = WorksheetFunction.Index(ws.Range(tkr_range), WorksheetFunction.Match(min_value, ws.Range(per_range), 0))
        ws.Cells(3, 16) = min_value
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 15) = WorksheetFunction.Index(ws.Range(tkr_range), WorksheetFunction.Match(max_volume, ws.Range(vol_range), 0))
        ws.Cells(4, 16) = max_volume
       
    Next ws

End Sub

