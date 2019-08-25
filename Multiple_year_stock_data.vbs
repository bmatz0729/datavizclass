Sub TicTock()
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    ws.Activate

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    Dim ticker As String
    Dim Open_price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim vol As Double
    vol = 0
    Dim counter As Integer
    Dim Row As Double
    Dim Column As Double
    Row = 2
    Column = 1



    'Add Heading
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Vol"
    

    counter = 2
    
    Open_price = Cells(2, Column + 2).Value
         '

    For i = 2 To LastRow

      

      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         


          ticker = Cells(i, Column).Value
          vol = vol + Cells(i, 7).Value
          Close_Price = Cells(i, Column + 5).Value
          Yearly_Change = Close_Price - Open_price
          Cells(Row, Column + 9).Value = Yearly_Change
          If (Open_price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
          
          Range("I" & counter).Value = ticker
          Range("J" & counter).Value = Yearly_Change
          Range("K" & counter).Value = Percent_Change
          Range("L" & counter).Value = vol

          counter = counter + 1

          vol = 0


      Else

          vol = vol + Cells(i, 7).Value


      End If


    Next i

Next ws


End Sub