Attribute VB_Name = "Module1"
Sub Stocks():

'Setting needed variables
    Dim i As Long
    Dim Volume As LongLong
    Dim TotalRows As Long
    Dim AnalysisRows As Long
    Dim j As Long
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim MaxIncrease As Double
    Dim MaxIncreaseT As String
    Dim MaxDecrease As Double
    Dim MaxDecreaseT As String
    Dim MaxVolume As LongLong
    Dim MaxVolumeT As String
    
    Dim ws As Worksheet

    For Each ws In Worksheets

    
    'Counting rows and creating headers
        TotalRows = Cells(Rows.Count, 1).End(xlUp).Row
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Volume = 0
        
        j = 2
        
        For i = 2 To TotalRows
        
    'If current i cell is different than the cell above it, save the year open value
            If (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
                
                YearOpen = ws.Cells(i, 3).Value
    'If current i cell is different than the cell below it, print the ticker in the analysis, add volume for that row
    'Print volume in analysis, record year end value, print the difference between year open and close
    'Print the % change from open to close. If either value was 0, print 0, color the cells based on increase or decrease
            ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                
                
                Volume = CLngLng(Volume) + CLngLng(ws.Cells(i, 7).Value)
                ws.Cells(j, 12).Value = Volume
                 
                YearClose = ws.Cells(i, 6).Value
                ws.Cells(j, 10).Value = YearClose - YearOpen
                
                If (YearClose = 0) Then
                    ws.Cells(j, 11).Value = 0
                
                ElseIf (YearOpen = 0) Then
                    ws.Cells(j, 11).Value = 0
                
                Else
                    ws.Cells(j, 11).Value = YearClose / YearOpen - 1
                
                End If
                
                If (ws.Cells(j, 11).Value > MaxIncrease) Then
                    MaxIncrease = ws.Cells(j, 11).Value
                    MaxIncreaseT = ws.Cells(i, 1).Value

                ElseIf (ws.Cells(j, 11).Value < MaxDecrease) Then
                    MaxDecrease = ws.Cells(j, 11).Value
                    MaxDecreaseT = ws.Cells(i, 1).Value

                ElseIf (Volume > MaxVolume) Then
                    MaxVolume = Volume
                    MaxVolumeT = ws.Cells(i, 1).Value

                End If
                
                
                ws.Cells(j, 11).NumberFormat = "0.00%"
    
                If (ws.Cells(j, 10).Value > 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                
                ElseIf (ws.Cells(j, 10).Value < 0) Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                
                End If
                
                j = j + 1
                Volume = 0
    
    'If current i cell is not different from the i below it, add the volume for that row
            Else
                
                Volume = CLngLng(Volume) + CLngLng(ws.Cells(i, 7).Value)
            
            End If
        Next i

    'Building the table
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greates Total Volume"
        
        ws.Range("P2") = MaxIncreaseT
        ws.Range("P3") = MaxDecreaseT
        ws.Range("P4") = MaxVolumeT
        
        ws.Range("Q2") = MaxIncrease
        ws.Range("Q3") = MaxDecrease
        ws.Range("Q4") = MaxVolume
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        MaxIncrease = 0
        MaxDecrease = 0
        MaxVolume = 0
    Next ws
        
End Sub

