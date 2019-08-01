Attribute VB_Name = "Module1"
Sub Easy_Stocks()

        Dim Worksheet As Worksheet
        Dim Results_WorkSheet As Boolean
        Summary = False
        
        For Each Worksheet In Worksheets

        Dim Ticker_Number As String
            
            Dim Total_Table As Long
            Total_Table = 2
            
            Dim Total_Ticker_Volume As Double
            Total_Ticker_Volume = 0
            
            Dim Lastrow As Long
            Dim x As Long
            
            Last_row = Worksheet.Cells(Rows.Count, 1).End(xlUp).Row

            If Summary Then
                Worksheet.Cells(9.1).Value = "Ticker Volume"
                Worksheet.Cells(10, 1).Value = "Total Stock Volume"
            Else
                Summary = True
            End If
            For x = 2 To Last_row
                If Worksheet.Cells(x + 1, 1).Value <> Worksheet.Cells(x, 1).Value Then
                    Ticker_Number = Worksheet.Cells(x, 1).Value
                    Total_Ticker_Volume = Total_Ticker_Volume + Worksheet.Cells(x, 7).Value
                    Worksheet.Range("I" & Total_Table).Value = Ticker_Number
                    Worksheet.Range("J" & Total_Table).Value = Total_Ticker_Volume
                    Total_Table = Total_Table + 1
                    Total_Ticker_Volume = 0
                Else
                    Total_Ticker_Volume = Total_Ticker_Volume + Worksheet.Cells(x, 7).Value
                End If
            Next x
         Next Worksheet
End Sub

