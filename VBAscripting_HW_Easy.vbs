Sub Easy()

For Each ws In Worksheets

Dim lastrow As Long

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


Dim Ticker As String

Dim Stock_Volume As Double

Stock_Volume = 0
 

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

 
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"


For i = 2 To lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    Ticker = ws.Cells(i, 1).Value

    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    ws.Range("I" & Summary_Table_Row).Value = Ticker

    ws.Range("J" & Summary_Table_Row).Value = Stock_Volume

    Summary_Table_Row = Summary_Table_Row + 1

    Stock_Volume = 0
    

    Else

    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    End If

Next i


ws.Columns("A:K").AutoFit

Next ws

End Sub

