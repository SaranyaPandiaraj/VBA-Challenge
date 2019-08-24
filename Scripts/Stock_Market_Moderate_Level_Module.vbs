
'Module For Stock Market Analysis

Sub Stock_Market_Medium()

    'Looping through all the worksheets
    
    For Each ws In Worksheets
     
       'Clearing the Contents of the Target Cells
       
       For i = 9 To 12
       
        ws.Columns(i).Clear
        
       Next i
       
       ws.Range("O1:Q4").Clear
        
       'Declaring Variables Used in this Module

        Dim LastRow1 As Long
        Dim TickerSymbol As String
        Dim TotalStockVolume As Double
        Dim RowNum As Integer
        Dim VolumeValue As Double
        
        Dim OpenValue As Double
        Dim CloseValue As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim NextVal As Long
        
        'Initializing the Initial Value
        
        TotalStockVolume = 0
        RowNum = 2
        NextVal = 2
    
        'Assigning the Header Values in each sheet
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        
        'Making the Headers in Bold
        
        ws.Range("I1:L1").Font.Bold = True
        
       'Auto Fit the Header Values
       
        ws.Columns("I:L").AutoFit
        
' --------------------------------------------- Part 1 : Easy Level -------------------------------------------------------- '
        
        'Identifying the Last Row Value in each sheet

        LastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Looping from 2 to the Last Row Value in each sheet
        
        For i = 2 To LastRow1
        
            'Initializing the Volume Value and Calculating the Sum of the Stock Value
            
            VolumeValue = ws.Cells(i, 7).Value
            
            TotalStockVolume = TotalStockVolume + VolumeValue
            
            'Condition for Validating the Same Ticker Symbol or Not
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Initialzing the Ticker Symbol Value
                
                TickerSymbol = ws.Cells(i, 1).Value
                
                'Displaying the Ticker Symbol and Total Stock Volume in the Summary Table (For each Ticker Symbol)
                
                ws.Range("I" & RowNum) = TickerSymbol
                
                ws.Range("L" & RowNum) = TotalStockVolume
                
                'Resetting the Total Stock value
                
                TotalStockVolume = 0

' --------------------------------------------- Part 2 : Moderate Level ------------------------------------------------------- '
                
                'Initializing the Open Value, Close Value and Calculating the Yearly Change
                
                OpenValue = ws.Range("C" & NextVal)
                CloseValue = ws.Range("F" & i)
                
                YearlyChange = CloseValue - OpenValue
                ws.Range("J" & RowNum).Value = YearlyChange
                
                'Conditional Formatting the YearlyChange Colum (positive change in green and negative change in red)
                
                If ws.Range("J" & RowNum).Value >= 0 Then
                    ws.Range("J" & RowNum).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & RowNum).Interior.ColorIndex = 3
                End If
                
                'Calculating Percent Change
                
                If OpenValue = 0 Then
                
                    PercentChange = 0
                    
                    Else
                    
                    PercentChange = YearlyChange / OpenValue
                    
                End If
                
                
                ws.Range("K" & RowNum).Value = PercentChange
                
                ' Formatting the Percent Change Column
                
                ws.Range("K" & RowNum).NumberFormat = "0.00%"
                
                    
                'Incrementing the Row Num and NextVal by 1
                
                RowNum = RowNum + 1
                
                NextVal = i + 1
                
            End If
        
        Next i
        
        'Adding Borders
        
        For i = 1 To Cells(Rows.Count, "I").End(xlUp).Row
        
            ws.Range("I" & i).BorderAround ColorIndex:=1, Weight:=xlThin
            ws.Range("J" & i).BorderAround ColorIndex:=1, Weight:=xlThin
            ws.Range("K" & i).BorderAround ColorIndex:=1, Weight:=xlThin
            ws.Range("L" & i).BorderAround ColorIndex:=1, Weight:=xlThin
        
        Next i
        
    Next ws
    
End Sub



