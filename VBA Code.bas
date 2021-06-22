Attribute VB_Name = "Module1"
Sub stock_ticker()
    
    For Each ws In Worksheets
        ' Set an initial variable for holding the ticker
        Dim Ticker_Name As String

        ' Set an initial variable for holding the total volume per ticker
        Dim Tot_Vol As Double
        Tot_Vol = 0

        ' Keep track of the location for each ticker in the summary table
    
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Lst_Rw = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Starting_Value = ws.Cells(2, 3).Value
    
        For i = 2 To Lst_Rw
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker name
      
                Ticker_Name = ws.Cells(i, 1).Value
                End_Value = ws.Cells(i, 6).Value
                Yearly_Change = End_Value - Starting_Value
                    If Starting_Value <> 0 Then
                        Percentage_Change = (Yearly_Change / Starting_Value)
                    Else
                        Percentage_Change = " "
                    End If
                Starting_Value = ws.Cells(i + 1, 3).Value

                ' Add to the Total Volume
                Tot_Vol = Tot_Vol + ws.Cells(i, 7).Value

                ' Print the Ticker in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                ws.Range("I1").Value = "Ticker Name"
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("J1").Value = "Yearly Change"
                ws.Range("K" & Summary_Table_Row).Value = Format(Percentage_Change, "0.00%")
                ws.Range("K1").Value = "Percentage Change"
                ws.Range("L" & Summary_Table_Row).Value = Tot_Vol
                ws.Range("L1").Value = "Total Volume"
        
                    If (ws.Range("J" & Summary_Table_Row).Value > 0) Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
      
      
      
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the
                Tot_Vol = 0

            ' If the cell immediately following a row is the same brand...
            Else

                ' Add to the Total Volume
                Tot_Vol = Tot_Vol + ws.Cells(i, 7).Value
                Starting_Value = Starting_Value
              

            End If

        Next i
       ws.Cells.Columns.AutoFit
    Next ws

End Sub

        







        





