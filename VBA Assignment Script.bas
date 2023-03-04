Attribute VB_Name = "Module1"
Sub Stocktotal()


'Sheets("A").Activate


For Each WS In Worksheets

       WS.Cells(1, 9).Value = "Ticker"
       WS.Cells(1, 10).Value = "Yearly Change"
       WS.Cells(1, 11).Value = "Yearly percentage"
       WS.Cells(1, 12).Value = "Total Stock Volume"
       
       
    Dim Ticker_Name As String
        Ticker_Name = 2
        
    Dim Next_Ticker_name As String
    
    Dim Vol_total As Double
        Vol_total = 0
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
    Dim yearOpen As Double
         yearOpen = 0
         
    Dim yearclose As Double
    
    Dim lastrow As Long
        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    yearOpen = WS.Cells(2, 3).Value
    

    For i = 2 To lastrow
    
        Ticker_Name = WS.Cells(i, 1).Value
        Next_Ticker_name = WS.Cells(i + 1, 1).Value
       
        
        
    If Next_Ticker_name <> Ticker_Name Then
    
        
    yearclose = WS.Cells(i, 6).Value
        
    
    WS.Cells(Summary_Table_Row, 9) = Ticker_Name
    
    Vol_total = Vol_total + WS.Cells(i, 7).Value
    WS.Cells(Summary_Table_Row, 12) = Vol_total
    
    YearChange = yearclose - yearOpen
    WS.Cells(Summary_Table_Row, 10) = YearChange
    
    PercentChange = (yearclose - yearOpen) / yearclose
    WS.Cells(Summary_Table_Row, 11) = PercentChange
    WS.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
    yearOpen = WS.Cells(i + 1, 3).Value
       
    Summary_Table_Row = Summary_Table_Row + 1
    Vol_total = 0
    
    Else
 
        Vol_total = Vol_total + WS.Cells(i, 7).Value
    End If
    
    Next i
    
    For i = 2 To lastrow
      
      If WS.Cells(i, 10).Value >= 0 Then
      WS.Cells(i, 10).Interior.ColorIndex = 4
      
      ElseIf WS.Cells(i, 10).Value < 0 Then
      WS.Cells(i, 10).Interior.ColorIndex = 3
      
      End If
      Next i
      
    WS.Range("P1").Value = "Ticker"
    WS.Range("Q1").Value = "Value"
    
    WS.Range("O2").Value = "Greatest % Increase"
    WS.Range("O3").Value = "Greatest%Decrease"
    WS.Range("O4").Value = "Greates Total Volume"
    
    Dim GI As Double
        GI = 0
    Dim GD As Double
        GD = 0
    Dim GV As Double
        GV = 0
    
    For i = 2 To lastrow
        If WS.Cells(i, 11).Value > GI Then
        GI = WS.Cells(i, 11).Value
        
        End If
         
        Next i
        
        WS.Range("Q2").Value = GI
        WS.Range("Q2").Style = "Percent"
        WS.Range("Q2").NumberFormat = "0.00%"
        WS.Range("P2").Value = WS.Cells(i, 9).Value
        
        
        
       
        
    For i = 2 To lastrow
        If WS.Cells(i, 11).Value < GD Then
        GD = WS.Cells(i, 11).Value
        End If
        Next i
        
        WS.Range("Q3").Value = GD
        WS.Range("Q3").Style = "Percent"
        WS.Range("Q3").NumberFormat = "0.00%"
        WS.Range("P3").Value = WS.Cells(i, 9).Value
        
        
        
      
        
    For i = 2 To lastrow
        If WS.Cells(i, 12).Value > GV Then
        GV = WS.Cells(i, 12).Value
       End If
        
        Next i
        
        WS.Range("Q4").Value = GV
        WS.Range("P4").Value = WS.Cells(i, 9).Value
        
        
        
        
        
    Next WS
    
        
    End Sub




