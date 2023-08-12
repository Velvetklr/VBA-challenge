Sub stock()
'constants
    Const CLOSED_COL As Integer = 6
    Const VOLUME_COL As Integer = 7

'Setting Dims
    Dim Opened As Variant
    Dim Closed As Variant
    Dim Vol As LongLong
    Dim Change As Variant
    Dim Percent As Variant
    Dim TickerTable As Integer
    Dim LastRow As Long
    Dim index As Long
    Dim MaxIncTick As String
    Dim MaxIncVal As Variant
    Dim MinIncTick As String
    Dim MinIncVal As Variant
    Dim MaxVolTick As String
    Dim MaxVolVal As LongLong
    Dim ws As Worksheet

'Run on each Workskeet
For Each ws In Worksheets
    ws.Activate
    
'Setting Variable
    Start = 2
    Vol = 0
    TickerTable = 2
    MaxIncVal = -99999
    MinIncVal = 0
   
'Get Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For index = 2 To LastRow
        Ticker = ws.Cells(index, 1).Value
        Vol = Vol + ws.Cells(index, VOLUME_COL).Value
        If ws.Cells(index + 1, 1).Value <> Ticker Then
        
            'Inputs
                If (ws.Cells(Start, 3) = 0) Then
                    For find_value = Start To index
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            Start = find_value
                    Exit For
                        End If
                    Next find_value
                End If
            Change = (ws.Cells(index, 6) - ws.Cells(Start, 3))
            Percent = Change / ws.Cells(Start, 3)
            
            If Percent > MaxIncVal Then
                MaxIncVal = Percent
                MaxIncTick = Ticker
            End If
            
            If Percent < MinIncVal Then
                MinIncVal = Percent
                MinIncTick = Ticker
            End If
            If Vol > MaxVolVal Then
                MaxVolVal = Vol
                MaxVolTick = Ticker
            End If
                
            Start = index + 1
            'Outputs
            ws.Range("I" & TickerTable).Value = Ticker
            ws.Range("L" & TickerTable).Value = Vol
            ws.Range("J" & TickerTable).Value = Change
            ws.Range("K" & TickerTable).Value = Percent
            ws.Range("K" & TickerTable).NumberFormat = "0.00%"
            Range("Q2").NumberFormat = "0.00%"
            Range("Q3").NumberFormat = "0.00%"
                           
                If Change > 0 Then
                    ws.Range("J" & TickerTable).Interior.ColorIndex = 4
                ElseIf Change < 0 Then
                    ws.Range("J" & TickerTable).Interior.ColorIndex = 3
                End If
                
                
            'Prepare for next Stock
            TickerTable = TickerTable + 1
            Vol = 0
            
        End If
    
    Next index
    'Set row headers
        ws.Range("I1,P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P2").Value = MaxIncTick
        ws.Range("Q2").Value = MaxIncVal
        ws.Range("P3").Value = MinIncTick
        ws.Range("Q3").Value = MinIncVal
        ws.Range("P4").Value = MaxVolTick
        ws.Range("Q4").Value = MaxVolVal
Next ws
End Sub


