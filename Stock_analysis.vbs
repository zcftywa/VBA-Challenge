
Sub Stock_analysis()
   
    Dim Ticker As String
    Dim LastRow As Long
    Dim Price_Open As Double
    Dim Price_Close As Double
    Dim Price_Change As Double
    Dim Price_PerctChg As Double
    Dim Stock_Vol As Double
    Dim Output_Row As Integer
    Dim Output_Col() As String
    Dim Column_Names() As String
    Dim Output_Array() As Variant
    Dim Great_Perct_Inc As Variant
    Dim Great_Perct_Dec As Variant
    Dim Great_Total_Vol As Variant
    Dim Final_Output_Header() As String

      
    Dim ws As Worksheet
     
   
    Great_Perct_Inc = Array("Greatest % Increase", Ticker, 0)
    Great_Perct_Dec = Array("Greatest % Decrease", Ticker, 0)
    Great_Total_Vol = Array("Greatest Total Volume", Ticker, 0)
        
        
   
    For Each ws In Worksheets
    
      
        Output_Col = Split("K,L,M,N", ",")
        Column_Names = Split("Ticker Symbol,Yearly Change $,Percent Change %,Total Stock Volume", ",")
        Output_Row = 1
        Stock_Vol = 0
        Price_Open = ws.Cells(2, 3).Value
        Ticker = ws.Cells(2, 1).Value
        Final_Output_Header = Split("Ticker,Value", ",")
        
      
        Great_Perct_Inc = Array("Greatest % Increase", Ticker, 0)
        Great_Perct_Dec = Array("Greatest % Decrease", Ticker, 0)
        Great_Total_Vol = Array("Greatest Total Volume", Ticker, 0)
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 0 To UBound(Output_Col)
            ws.Range(Output_Col(i) & Output_Row).Value = Column_Names(i)
        Next i
        
  
        Output_Row = Output_Row + 1
        
        
       
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                
                Ticker = ws.Cells(i, 1).Value
                
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
             
                Price_Close = ws.Cells(i, 6).Value
                
                Price_Change = Price_Close - Price_Open
                
                If Price_Open <> 0 Then
                    Price_PerctChg = Price_Change / Price_Open
                Else
                    Price_PerctChg = Price_Change / 0.01
                End If
                
                If Price_PerctChg > Great_Perct_Inc(2) Then
                    Great_Perct_Inc(1) = Ticker
                    Great_Perct_Inc(2) = Price_PerctChg
                ElseIf Price_PerctChg < Great_Perct_Dec(2) Then
                    Great_Perct_Dec(1) = Ticker
                    Great_Perct_Dec(2) = Price_PerctChg
                End If
                

                If Stock_Vol > Great_Total_Vol(2) Then
                    Great_Total_Vol(1) = Ticker
                    Great_Total_Vol(2) = Stock_Vol
                End If
                
                
                
                Output_Array = Array(Ticker, Price_Change, Price_PerctChg, Stock_Vol)
                
                For k = 0 To UBound(Output_Col)
                    If k = 1 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        If Output_Array(k) > 0 Then
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 4
                        ElseIf Output_Array(k) < 0 Then
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 3
                        Else
                            ws.Range(Output_Col(k) & Output_Row).Interior.ColorIndex = 6
                        End If
                    ElseIf k = 2 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        ws.Range(Output_Col(k) & Output_Row).NumberFormat = "#,###.00%"
                    ElseIf k = 3 Then
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                        ws.Range(Output_Col(k) & Output_Row).NumberFormat = "#,###"
                    Else
                        ws.Range(Output_Col(k) & Output_Row).Value = Output_Array(k)
                    End If
                Next k
            
                Ticker = ws.Cells(i + 1, 1).Value
                Stock_Vol = 0
                Price_Open = ws.Cells(i + 1, 3).Value
                Output_Row = Output_Row + 1
                            
         
            Else
                
                Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
 
        ws.Range("R1:S1").Value = Final_Output_Header
        ws.Range("Q2:S2").Value = Great_Perct_Inc
        ws.Range("S2").NumberFormat = "#,###.00%"
        ws.Range("Q3:S3").Value = Great_Perct_Dec
        ws.Range("S3").NumberFormat = "#,###.00%"
        ws.Range("Q4:S4").Value = Great_Total_Vol
        ws.Range("S4").NumberFormat = "#,###"
        
        ws.Columns("A:Q").EntireColumn.AutoFit
    Next

    
    
End Sub

