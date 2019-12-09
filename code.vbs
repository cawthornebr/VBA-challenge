Sub TickerHw()
'create variable for the ticker
Dim Ticker As String

'create variable for open line and open price for each set of tickers. set open line to 2 to avoid headers
Dim Ol As Long
Dim Op As Double
Ol = 2

'create variable for close line and open price for each set of tickers
Dim Cl As Long
Dim Cp As Double

'create variable for greatest % increase/decrease and gratest total volume. for each we need a value, name, and starting value. greatest total volume needs to be set to "Variant" becaues number exceeds long
Dim Giv As Double
Dim Gin As String
Giv = 0
Dim Gdv As Double
Dim Gdn As String
Gdv = 0
Dim Gtv As Variant
Dim Gtn as String
Gtv = 0

'create variable for counting all new rows that are being created for sub totals. set to 2 to avoid headers
Dim count As Integer
count = 2

'loop through each sheet in the workbook
For Each ws In Worksheets

    'set current sheet to active
    ws.Activate

    'create variable for last row
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.count, "A").End(xlUp).Row

    'create new headers
    ws.Range("I1, P1").Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(2, 17).Select
    Selection.NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 17).Select
    Selection.NumberFormat = "0.00%"
    ws.Cells(4, 15).Value = "Greatest total volume"

    'change format for column k to percent
    Columns("K:K").Select
    Selection.NumberFormat = "0.00%"

    'change format for column g and l to number
    ActiveSheet.Columns("G:G").NumberFormat = "0"
    ActiveSheet.Columns("L:L").NumberFormat = "0"

    'loop through each row on current sheet
    For i = 2 To LastRow

        'find the last rown for each ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

            'set close line
            Cl = i

            'set open price
            Op = ws.Cells(Ol, 3).Value

            'set close price
            Cp = ws.Cells(Cl, 6).Value

            'add ticker label to column I
            ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
            
            If Op <> 0 Then

                'add yearly change value to column J
                ws.Cells(count, 10).Value = Cp - Op

                'set color to cell: red for negative values and green for positive values
                If ws.Cells(count, 10).Value < 0 Then
                    ws.Cells(count, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                End If

                'add percent change value to column K
                ws.Cells(count, 11).Value = (Cp / Op) - 1

            Else

                'if Op value is 0, then set yearly and percent change values to whatever the close value is and make cell green
                ws.Cells(count, 10).Value = Cp
                ws.Cells(count, 10).Interior.ColorIndex = 4
                ws.Cells(count, 11).Value = Cp

            End If

            'add total stock volume to column L
            ws.Cells(count, 12).Value = Application.Sum(ws.Range("G" & Ol & ":G" & Cl))

            'get the greatest total % increase value and ticker
            If ws.Cells(count, 11).Value > Giv Then
                Giv=ws.Cells(count, 11).Value
                Gin=ws.Cells(count,9).Value
            End If

            'get the greatest total % decrease value and ticker
            If ws.Cells(count, 11).Value < Gdv Then
                Gdv=ws.Cells(count, 11).Value
                Gdn=ws.Cells(count,9).Value
            End If

            'get the greatest total volume ticker and value
            If ws.Cells(count, 12).Value > Gtv Then
                Gtv=ws.Cells(count, 12).Value
                Gtn=ws.Cells(count,9).Value
            End If

            'update counter by one so it starts one row down for the new ticker
            count = count + 1

            'Reset open line for new ticker
            Ol = i + 1

        End If

    Next i
    
    'rest the open line and count to 2
    Ol = 2
    count = 2

    'add the greatest increase value and ticker
    ws.Cells(2,17).Value = Giv
    Giv = 0
    ws.Cells(2,16).Value = Gin

    'add the greatest decrease value and ticker
    ws.Cells(3,17).Value = Gdv
    Gdv = 0
    ws.Cells(3,16).Value = Gdn

    'add the greatest decrease value and ticker
    ws.Cells(4,17).Value = Gtv
    Gtv = 0
    ws.Cells(4,16).Value = Gtn

Next ws

End Sub
