VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub QuarterlyStockAnalysis_ConsolidatedToQ1()
    Dim ws As Worksheet
    Dim wsQ1 As Worksheet
    Dim lastRow As Long, outputRow As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double
    Dim totalVolume As Double
    Dim currentQuarter As String, prevQuarter As String
    Dim i As Long
    Dim rawDate As Double
    Dim formattedDate As Date
    Dim yearVal As Integer, quarterVal As Integer
    
    ' Variables for greatest values
    Dim maxIncrease As Double, maxDecrease As Double, maxVolume As Double
    Dim maxIncreaseTicker As String, maxDecreaseTicker As String, maxVolumeTicker As String
    
    ' Initialize variables
    maxIncrease = -99999
    maxDecrease = 99999
    maxVolume = 0

    ' Set the Q1 sheet to store the final results
    On Error Resume Next
    Set wsQ1 = ThisWorkbook.Sheets("Q1")
    On Error GoTo 0
    
    ' Check if Q1 sheet exists
    If wsQ1 Is Nothing Then
        MsgBox "Q1 sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Clear previous output in Q1 sheet
    wsQ1.Cells(1, 9).Resize(wsQ1.Rows.Count, 5).ClearContents
    
    ' Add headers for output in columns 9 to 13 (I to M)
    wsQ1.Cells(1, 9).Value = "Ticker"
    wsQ1.Cells(1, 10).Value = "Quarter"
    wsQ1.Cells(1, 11).Value = "Quarterly Change"
    wsQ1.Cells(1, 12).Value = "Percentage Change"
    wsQ1.Cells(1, 13).Value = "Total Volume"
    
    ' Add headers for greatest values in columns 16 to 18 (P to R)
    wsQ1.Cells(1, 16).Value = "Criteria"
    wsQ1.Cells(1, 17).Value = "Ticker"
    wsQ1.Cells(1, 18).Value = "Value"
    
    wsQ1.Cells(2, 16).Value = "Greatest % Increase"
    wsQ1.Cells(3, 16).Value = "Greatest % Decrease"
    wsQ1.Cells(4, 16).Value = "Greatest Total Volume"
    
    ' Initialize output row in Q1 sheet
    outputRow = 2

    ' Loop through each sheet Q1 to Q4
    For Each ws In ThisWorkbook.Sheets
        If ws.Name Like "Q*" And ws.Name <> "Q1" Then  ' Process Q2, Q3, and Q4 sheets
            ' Reset variables for each sheet
            maxIncrease = -99999
            maxDecrease = 99999
            maxVolume = 0
            totalVolume = 0
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Loop through the data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 1).Value
                rawDate = ws.Cells(i, 2).Value
                ' Convert YYYYMMDD to actual date
                formattedDate = DateSerial(Left(rawDate, 4), Mid(rawDate, 5, 2), Right(rawDate, 2))
                
                ' Extract Year and Quarter
                yearVal = Year(formattedDate)
                quarterVal = Application.WorksheetFunction.RoundUp(Month(formattedDate) / 3, 0)
                currentQuarter = yearVal & " Q" & quarterVal
                
                ' Detect new ticker or new quarter
                If i = 2 Or ticker <> ws.Cells(i - 1, 1).Value Or currentQuarter <> prevQuarter Then
                    openPrice = ws.Cells(i, 3).Value  ' Column C for Opening Price
                    totalVolume = 0
                End If
                
                ' Accumulate volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value ' Column G for Volume
                
                ' Detect end of quarter or ticker
                If i = lastRow Or ticker <> ws.Cells(i + 1, 1).Value Then
                    closePrice = ws.Cells(i, 6).Value ' Column F for Closing Price
                    
                    ' Calculate changes
                    Dim quarterlyChange As Double, percentageChange As Double
                    quarterlyChange = closePrice - openPrice
                    
                    If openPrice <> 0 Then
                        percentageChange = (quarterlyChange / openPrice) * 100
                    Else
                        percentageChange = 0
                    End If
                    
                    ' Output results to columns 9-13 (I-M) in Q1 sheet
                    wsQ1.Cells(outputRow, 9).Value = ticker
                    wsQ1.Cells(outputRow, 10).Value = currentQuarter
                    wsQ1.Cells(outputRow, 11).Value = quarterlyChange
                    wsQ1.Cells(outputRow, 12).Value = percentageChange
                    wsQ1.Cells(outputRow, 13).Value = totalVolume
                    
                    ' Check for greatest % increase
                    If percentageChange > maxIncrease Then
                        maxIncrease = percentageChange
                        maxIncreaseTicker = ticker
                    End If
                    
                    ' Check for greatest % decrease
                    If percentageChange < maxDecrease Then
                        maxDecrease = percentageChange
                        maxDecreaseTicker = ticker
                    End If
                    
                    ' Check for greatest total volume
                    If totalVolume > maxVolume Then
                        maxVolume = totalVolume
                        maxVolumeTicker = ticker
                    End If
                    
                    ' Move to the next output row
                    outputRow = outputRow + 1
                End If
                
                ' Store previous quarter value
                prevQuarter = currentQuarter
            Next i
        End If
    Next ws
    
    ' Output greatest values to Q1 sheet (Columns P-R)
    wsQ1.Cells(2, 17).Value = maxIncreaseTicker
    wsQ1.Cells(2, 18).Value = maxIncrease
    wsQ1.Cells(3, 17).Value = maxDecreaseTicker
    wsQ1.Cells(3, 18).Value = maxDecrease
    wsQ1.Cells(4, 17).Value = maxVolumeTicker
    wsQ1.Cells(4, 18).Value = maxVolume
    
    ' Format percentage column in Q1
    wsQ1.Columns(12).NumberFormat = "0.00%"
    
    MsgBox "Quarterly stock analysis completed for all quarters and consolidated in Q1 sheet!", vbInformation
End Sub




