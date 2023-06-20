Attribute VB_Name = "VBA_Challenge"
Sub Get_Ticker():
Dim ws As Worksheet
For Each ws In Worksheets

        
'Define Variables
Dim ticker As String
Dim mar_open As Double
Dim mar_close As Double
Dim totalvolume As Double
Dim x As Double
Dim y As Double
Dim LastRow As Long
Dim maxpercent As Double
Dim maxpercentticker As String
Dim maxvolume As Double
Dim maxvolticker As String
Dim minpercent As Double
Dim minpercentticker As String
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
x = 2
y = 2


'Format Output area
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Columns("A:Q").AutoFit

'for loops
For i = 2 To LastRow
    
    'conditional parameters
    'checking the next cell with the current
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'conditional instructions
        
        totalvolume = totalvolume + ws.Cells(i, 7)
        ticker = ws.Cells(i, 1).Value

        
        mar_open = ws.Cells(x, 3).Value
        mar_close = ws.Cells(i, 6).Value
        
        'Assign Outputs
        'perform calculations
        Yearlychange = (mar_close - mar_open)
        percentchange = Yearlychange / mar_open
        ws.Range("I" & y).Value = ticker
        ws.Range("J" & y).Value = Yearlychange
            If ws.Range("J" & y).Value > 0 Then
            ws.Range("J" & y).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & y).Value < 0 Then
            ws.Range("J" & y).Interior.ColorIndex = 3
            End If
        'Formatting new columns for readability
        ws.Range("J" & y).Value = Yearlychange
        ws.Range("J" & y).NumberFormat = "$0.00"
        ws.Range("K" & y).Value = percentchange
        ws.Range("K" & y).NumberFormat = "0.00%"
        ws.Range("L" & y).Value = totalvolume
        ws.Range("L" & y).NumberFormat = "#,##0"
        totalvolume = 0
        y = y + 1
        x = i + 1
    
    'add to totalvolume
    Else
        totalvolume = totalvolume + ws.Cells(i, 7)

    End If



Next i




maxvolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow).Value)
maxvolticker = 1 + WorksheetFunction.Match(maxvolume, (ws.Range("L2:L" & LastRow).Value), 0)
ws.Range("q4").Value = maxvolume
ws.Range("P4") = ws.Cells(maxvolticker, 9)

maxpercent = WorksheetFunction.Max(ws.Range("K2:K" & LastRow).Value)
maxpercentticker = 1 + WorksheetFunction.Match(maxpercent, (ws.Range("K2:K" & LastRow).Value), 0)
ws.Range("q2").Value = maxpercent
ws.Range("q2").NumberFormat = "0.00%"
ws.Range("P2") = ws.Cells(maxpercentticker, 9)

minpercent = WorksheetFunction.Min(ws.Range("K2:K" & LastRow).Value)
minpercentticker = 1 + WorksheetFunction.Match(minpercent, (ws.Range("K2:K" & LastRow).Value), 0)
ws.Range("q3").Value = minpercent
ws.Range("q3").NumberFormat = "0.00%"
ws.Range("P3") = ws.Cells(minpercentticker, 9)

Next ws



End Sub

Sub clear_contents():

Range("K2:N2").ClearContents

End Sub
