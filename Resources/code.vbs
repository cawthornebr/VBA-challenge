Sub TickerHw()
'create variable for the ticker
Dim Ticker As String

'create variable for open line, open price
Dim Ol As Long
Dim Op As Double
'open line starts at 2 to skip the headers row
Ol = 2

'create variable for close line, close price
Dim Cl As Long
Dim Cp As Double

'create variable for counting all new rows
Dim count As Integer
count = 2
For Each ws In Worksheets
    ws.Activate
    'Finding last row with the variable LastRow
    '---------
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.count, "A").End(xlUp).Row - 1
    '---------

    'Create new headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    Columns("K:K").Select
    Selection.Style = "Percent"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    'set column g to number format so we can add for total stock volume
    ActiveSheet.Columns("G:G").NumberFormat = "0"
    ActiveSheet.Range("A1").Select

    'start loop to go through all rows
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            'Set closing line as variablea
            Cl = i
            Op = ws.Cells(Ol, 3).Value
            Cp = ws.Cells(Cl, 6).Value
            'add ticker label to column I
            ws.Cells(count, 9).Value = ws.Cells(i, 1).Value
            If Cp <> 0 Then
                ws.Cells(count, 10).Value = Cp - Op
                If ws.Cells(count, 10).Value < 0 Then
                    ws.Cells(count, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(count, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(count, 11).Value = Op / Cp
                ws.Cells(count, 12).Value = Application.Sum(ws.Range("G" & Ol & ":G" & Cl))
            Else
                ws.Cells(count, 10).Value = "0"
                ws.Cells(count, 11).Value = "0"
            End If
            'update counter by one so it starts one row down for the new ticker
            count = count + 1
            'format colors for column J

            'Reset open line for new ticker
            Ol = i + 1
        Else
        End If
    Next i
    Ol = 2
    count = 2
Next ws
End Sub