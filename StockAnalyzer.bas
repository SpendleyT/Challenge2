Attribute VB_Name = "StockAnalyzer"
Sub Analyze_Sheet()

    Dim i As Integer, j As Integer
    Dim currSheet As String
    Dim currSymbol As String, newSymbol As String
    Dim currOpen As Double, currClose As Double
    Dim totalShares As LongLong
    Dim rowCount As Long
    Dim tableCount As Integer
    Dim wSheet As Worksheet
    Dim maxChange As Double, minChange As Double
    Dim maxTicker As String, minTicker As String, volTicker As String
    Dim maxVolume As LongLong
    
    'Loop thru sheets in workbook
    For Each wSheet In ActiveWorkbook.Worksheets
        
        wSheet.Activate
        
        'Format headers, labels and columns
        Table_Setup wSheet
        Format_Columns wSheet
                        
        'Set first table row
        tableCount = 2
    
        'Get row count
        LastRow = ActiveSheet.Cells(ActiveSheet.Cells.Rows.Count, "C").End(xlUp).Row
        rowCount = WorksheetFunction.CountA(ActiveSheet.Range("C1:C" & LastRow))
     
        'Process first symbol row
        ActiveSheet.Range("A2").Select
        currSymbol = ActiveCell.Value
        currOpen = ActiveCell.Offset(0, 2).Value
        currClose = ActiveCell.Offset(0, 5).Value
        totalShares = ActiveCell.Offset(0, 6).Value
        
        'Loop thru rows in sheet
        For j = 3 To rowCount + 1
            'Repeat for each row
            Cells(j, 1).Select
            newSymbol = ActiveCell.Value
            
            'check for new ticker
            If newSymbol = currSymbol Then
                currClose = ActiveCell.Offset(0, 5).Value
                totalShares = totalShares + ActiveCell.Offset(0, 6).Value
            Else
                'Update table
                ActiveSheet.Cells(tableCount, 9).Value = currSymbol
                ActiveSheet.Cells(tableCount, 10).Value = currClose - currOpen
                ActiveSheet.Cells(tableCount, 11).Value = FormatPercent((currClose - currOpen) / currOpen, 2)
                ActiveSheet.Cells(tableCount, 12).Value = totalShares
                
                'Color code change column
                If currClose - currOpen > 0 Then
                    ActiveSheet.Cells(tableCount, 10).Interior.Color = vbGreen
                ElseIf currClose - currOpen < 0 Then
                    ActiveSheet.Cells(tableCount, 10).Interior.Color = vbRed
                End If
                
                'Update max/min if necessary
                If ActiveSheet.Cells(tableCount, 11).Value > maxChange Then
                    maxChange = ActiveSheet.Cells(tableCount, 11).Value
                    maxTicker = ActiveCell.Value
                ElseIf ActiveSheet.Cells(tableCount, 11).Value < minChange Then
                    minChange = ActiveSheet.Cells(tableCount, 11).Value
                    minTicker = ActiveCell.Value
                End If
                
                'Update max value if necessary
                If ActiveSheet.Cells(tableCount, 12).Value > maxVolume Then
                    maxVolume = ActiveSheet.Cells(tableCount, 12).Value
                    volTicker = ActiveCell.Value
                End If
                
                'start new ticket count
                currSymbol = newSymbol
                currOpen = ActiveCell.Offset(0, 2).Value
                currClose = ActiveCell.Offset(0, 5).Value
                totalShares = ActiveCell.Offset(0, 6).Value
                
                tableCount = tableCount + 1
            End If
        Next j
        
        'Update max/min table
        Range("P2").Value = maxTicker
        Range("P3").Value = minTicker
        Range("P4").Value = volTicker
        
        Range("Q2").Value = FormatPercent(maxChange, 2)
        Range("Q3").Value = FormatPercent(minChange, 2)
        Range("Q4").Value = maxVolume
        
        'Clear max/min variables
        maxTicker = ""
        minTicker = ""
        volTicker = ""
        maxChange = 0
        minChange = 0
        maxVolume = 0
        Range("H1").Select
    Next

End Sub

Sub Table_Setup(wkSheet As Worksheet)

    With wkSheet
        'Format Headers
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        
        'Format max/min change table
        .Range("O2").Value = "Greatest % Increase"
        .Range("O3").Value = "Greatest % Decrease"
        .Range("O4").Value = "Greatest Total Volume"
        
        .Range("P1").Value = "Ticker"
        .Range("Q1").Value = "Value"
    End With
    
End Sub

Sub Format_Columns(ws As Worksheet)

    With ws
        Range("J:J").ColumnWidth = 12.11
        Range("K:K").ColumnWidth = 13.11
        Range("L:L").ColumnWidth = 19.11
        Range("O:O").ColumnWidth = 19.11
        Range("Q:Q").ColumnWidth = 12.11
    End With
    
End Sub
