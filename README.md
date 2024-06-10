# VBA-ChallengeWorksheets

("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set summarySheet = ThisWorkbook.Worksheets.Add
    summarySheet.Name = "Summary"
    
    ' Set the headers for the summary sheet
    With summarySheet
        .Cells(1, 1).Value = "Ticker Symbol"
        .Cells(1, 2).Value = "Volume of Stock"
        .Cells(1, 3).Value = "Open Price"
        .Cells(1, 4).Value = "Close Price"
    End With
    
    summaryRow = 2
    
    ' Loop through each worksheet name
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row and copy the data to the summary sheet
        For j = 2 To lastRow
            summarySheet.Cells(summaryRow, 1).Value = ws.Cells(j, "A").Value ' Ticker Symbol
            summarySheet.Cells(summaryRow, 2).Value = ws.Cells(j, "B").Value ' Volume of Stock
            summarySheet.Cells(summaryRow, 3).Value = ws.Cells(j, "C").Value ' Open Price
            summarySheet.Cells(summaryRow, 4).Value = ws.Cells(j, "D").Value ' Close Price
            summaryRow = summaryRow + 1
        Next j
    Next i
End Sub


Sub ConsolidateStockData()
    Dim wsNames As Variant
    Dim summarySheet As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim j As Long
    
    ' List of worksheet names to process
    wsNames = Array("A", "B", "C", "D", "E", "F")
    
    ' Create a new worksheet for the summary
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set summarySheet = ThisWorkbook.Worksheets.Add
    summarySheet.Name = "Summary"
    
    ' Set the headers for the summary sheet
    With summarySheet
        .Cells(1, 1).Value = "Ticker Symbol"
        .Cells(1, 2).Value = "Total Stock Volume"
        .Cells(1, 3).Value = "Quarterly Change"
        .Cells(1, 4).Value = "Percent Change"
    End With
    
    summaryRow = 2
    
    ' Loop through each worksheet name
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row and copy the data to the summary sheet
        For j = 2 To lastRow
            summarySheet.Cells(summaryRow, 1).Value = ws.Cells(j, "A").Value ' Ticker Symbol
            summarySheet.Cells(summaryRow, 2).Value = ws.Cells(j, "B").Value ' Total Stock Volume
            summarySheet.Cells(summaryRow, 3).Value = ws.Cells(j, "C").Value ' Quarterly Change
            summarySheet.Cells(summaryRow, 4).Value = ws.Cells(j, "D").Value ' Percent Change
            summaryRow = summaryRow + 1
        Next j
    Next i
End Sub

Sub AddHeadersAndConsolidateData()
    Dim wsNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim headerRow As Long

    ' List of worksheet names to process
    wsNames = Array("A", "B", "C", "D", "E", "F")
    
    ' Loop through each worksheet name
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
        
        ' Add headers in columns G, H, I, J
        headerRow = 1
        ws.Cells(headerRow, "G").Value = "Ticker Symbol"
        ws.Cells(headerRow, "H").Value = "Total Stock Volume"
        ws.Cells(headerRow, "I").Value = "Quarterly Change"
        ws.Cells(headerRow, "J").Value = "Percent Change"
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row and copy the data to the new columns
        For j = 2 To lastRow
            ws.Cells(j, "G").Value = ws.Cells(j, "A").Value ' Ticker Symbol
            ws.Cells(j, "H").Value = ws.Cells(j, "B").Value ' Total Stock Volume
            ws.Cells(j, "I").Value = ws.Cells(j, "C").Value ' Quarterly Change
            ws.Cells(j, "J").Value = ws.Cells(j, "D").Value ' Percent Change
        Next j
    Next i
End Sub

Sub AddHeadersAndCopyData()
    Dim wsNames As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim headerRow As Long

    ' List of worksheet names to process
    wsNames = Array("A", "B", "C", "D", "E", "F")
    
    ' Loop through each worksheet name
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
        
        ' Add headers in columns G, H, I, J
        headerRow = 1
        ws.Cells(headerRow, "G").Value = "Ticker Symbol"
        ws.Cells(headerRow, "H").Value = "Volume of Stock"
        ws.Cells(headerRow, "I").Value = "Open Price"
        ws.Cells(headerRow, "J").Value = "Close Price"
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row and copy the data to the new columns
        For j = 2 To lastRow
            ws.Cells(j, "G").Value = ws.Cells(j, "A").Value ' Ticker Symbol
            ws.Cells(j, "H").Value = ws.Cells(j, "B").Value ' Volume of Stock
            ws.Cells(j, "I").Value = ws.Cells(j, "C").Value ' Open Price
            ws.Cells(j, "J").Value = ws.Cells(j, "D").Value ' Close Price
        Next j
    Next i
End Sub

Sub ConsolidateStockData()
    Dim wsNames As Variant
    Dim summarySheet As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim j As Long
    
    ' List of worksheet names to process
    wsNames = Array("A", "B", "C", "D", "E", "F")
    
    ' Create a new worksheet for the summary
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set summarySheet = ThisWorkbook.Worksheets.Add
    summarySheet.Name = "Summary"
    
    ' Set the headers for the summary sheet
    With summarySheet
        .Cells(1, 1).Value = "Ticker Symbol"
        .Cells(1, 2).Value = "Total Stock Volume"
        .Cells(1, 3).Value = "Quarterly Change"
        .Cells(1, 4).Value = "Percent Change"
    End With
    
    summaryRow = 2
    
    ' Loop through each worksheet name
    For i = LBound(wsNames) To UBound(wsNames)
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
        
        ' Find the last row in the worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through each row and copy the data to the summary sheet
        For j = 2 To lastRow
            summarySheet.Cells(summaryRow, 1).Value = ws.Cells(j, "A").Value ' Ticker Symbol
            summarySheet.Cells(summaryRow, 2).Value = ws.Cells(j, "B").Value ' Total Stock Volume
            summarySheet.Cells(summaryRow, 3).Value = ws.Cells(j, "C").Value ' Quarterly Change
            summarySheet.Cells(summaryRow, 4).Value = ws.Cells(j, "D").Value ' Percent Change
            summaryRow = summaryRow + 1
        Next j
    Next i
End Sub
