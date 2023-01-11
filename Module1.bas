Attribute VB_Name = "Module1"
Sub Stock_data()

'retrieval of Data
Dim ticker As String
Dim closeP As Variant
Dim openP As Variant
Dim voltotal As Double
voltotal = 0

Dim ws As Worksheet

'summary table rows
Dim r As Integer
r = 2

'summary table2 variables
Dim greatin As Variant
Dim greatde As Variant
Dim greatvol As Double


For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("K:K").Style = "Percent"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest Percent Increase"
        ws.Range("O3").Value = "Greatest Percent Decrease"
        ws.Range("Q2:Q3").Style = "Percent"
        ws.Range("O4").Value = "Greatest Total Volume"
        
            
            For i = 2 To lastrow
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value And ws.Cells(i + 1, 1).Value = ws.Cells(i, 1) Then
                    openP = ws.Cells(i, 3).Value
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    closeP = ws.Cells(i, 6).Value
                    Change = closeP - openP
                    
                    ws.Cells(r, 9).Value = ticker
                    ws.Cells(r, 10).Value = Change
                    ws.Cells(r, 11).Value = Change / openP
                    ws.Cells(r, 12).Value = voltotal
                    r = r + 1
                    Change = 0
                    voltotal = 0
                Else
                    voltotal = voltotal + ws.Cells(i, 7).Value
                End If
            Next i
            
            lastrowx = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            For x = 2 To lastrowx
                If ws.Cells(x, 10).Value >= 0 Then
                    ws.Cells(x, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(x, 10).Interior.ColorIndex = 3
                End If
            Next x
            
            For j = 2 To lastrowx
                If ws.Cells(j + 1, 11).Value > greatinc Then
                    greatinc = ws.Cells(j + 1, 11).Value
                    ws.Range("P2").Value = ws.Cells(j + 1, 9).Value
                    ws.Range("Q2").Value = greatinc
                    
                ElseIf ws.Cells(j + 1, 11).Value < greatde Then
                    greatde = ws.Cells(j + 1, 11).Value
                    ws.Range("P3").Value = ws.Cells(j + 1, 9).Value
                    ws.Range("Q3").Value = greatde
                    
                ElseIf ws.Cells(j + 1, 12).Value > greatvol Then
                    greatvol = ws.Cells(j + 1, 12).Value
                    ws.Range("P4").Value = ws.Cells(j + 1, 9).Value
                    ws.Range("Q4").Value = greatvol
                End If
            Next j
            
            r = 2
            greatinc = 0
            greatde = 0
            greatvol = 0
            
        Next ws
End Sub
