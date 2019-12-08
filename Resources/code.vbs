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

'create variable for yearly change
'Dim Yc As Long

'create variable for counting all new rows
Dim count As Integer
count = 2

'Finding last row with the variable LastRow
'---------
Dim LastRow As Long
LastRow = ActiveSheet.Range("A" & Rows.count).End(xlUp).Row
'---------

'Create new headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Columns("K:K").Select
Selection.Style = "Percent"
Cells(1, 12).Value = "Total Stock Volume"
Columns("G:G").NumberFormat = "0"
Range("A1").Select

'start loop to go through all rows
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        'Set closing line as variablea
        Cl = i
        Op = Cells(Ol, 3).Value
        Cp = Cells(Cl, 6).Value
            'add ticker label to column I
        Cells(count, 9).Value = Cells(i, 1).Value
        If Cp <> 0 Then
            Cells(count, 10).Value = Cp - Op
            Cells(count, 11).Value = Op / Cp
            Cells(count, 12).Value = Application.Sum(Range("G" & Ol & ":G" & Cl))
        Else
            Cells(count, 10).Value = "0"
            Cells(count, 11).Value = "0"
        End If
        'update counter by one so it starts one row down for the new ticker
        count = count + 1
        
        'Reset open line for new ticker
        Ol = i + 1
    Else
    End If
Next i
End Sub