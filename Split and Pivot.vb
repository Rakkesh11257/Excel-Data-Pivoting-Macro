Sub SplitAndPivotValues()
    Dim ws As Worksheet
    Dim resultWs As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, k As Long
    Dim id As String
    Dim valuesB As Variant
    Dim valuesC As Variant
    Dim valuesD As Variant
    Dim valuesE As Variant
    Dim valuesF As Variant
    Dim valuesG As Variant
    Dim valuesH As Variant
    Dim valuesI As Variant
    Dim valuesJ As Variant
    Dim valuesK As Variant
    Dim valuesL As Variant
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your actual sheet name
    
    ' Create or reference the "Result" sheet
    On Error Resume Next
    Set resultWs = ThisWorkbook.Sheets("Result")
    On Error GoTo 0
    
    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
        resultWs.Name = "Result"
    End If
    
    ' Clear existing data in the "Result" sheet
    resultWs.Cells.Clear
    
    ' Add headers to the "Result" sheet
    resultWs.Cells(1, 1).Value = "ColumnA Header"
    resultWs.Cells(1, 2).Value = "ColumnB Header"
    resultWs.Cells(1, 3).Value = "ColumnC Header"
    resultWs.Cells(1, 4).Value = "ColumnD Header"
    resultWs.Cells(1, 5).Value = "ColumnE Header"
    resultWs.Cells(1, 6).Value = "ColumnF Header"
    resultWs.Cells(1, 7).Value = "ColumnG Header"
    resultWs.Cells(1, 8).Value = "ColumnH Header"
    resultWs.Cells(1, 9).Value = "ColumnI Header"
    resultWs.Cells(1, 10).Value = "ColumnJ Header"
    resultWs.Cells(1, 11).Value = "ColumnK Header"
    
    ' Set headers to bold
    resultWs.Rows(1).Font.Bold = True
    
    ' Highlight headers in yellow color
    resultWs.Rows(1).Interior.Color = RGB(255, 255, 0)
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Initialize the result row counter
    k = 2
    
    ' Loop through each row in column A until the last row with data
    For i = 2 To lastRow ' Start from row 2 to skip header
        ' Get ID from column A
        id = ws.Cells(i, 1).Value
        
        ' Get values from column B and split by semicolon
        valuesB = Split(ws.Cells(i, 2).Value, ";")
        
        ' Loop through each value and populate columns A and B accordingly in "Result" sheet
        For j = 0 To UBound(valuesB)
            ' Populate columns A and B in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 2).Value = Trim(valuesB(j))
        Next j
        
        ' Get values from column C and split by semicolon
        valuesC = Split(ws.Cells(i, 3).Value, ";")
        
        ' Loop through each value and populate columns A and C accordingly in "Result" sheet
        For j = 0 To UBound(valuesC)
            ' Populate columns A and C in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 3).Value = Trim(valuesC(j))
        Next j
        
        ' Repeat the above steps for columns D to L
        ' Get values from column D and split by semicolon
        valuesD = Split(ws.Cells(i, 4).Value, ";")
        
        ' Loop through each value and populate columns A and D accordingly in "Result" sheet
        For j = 0 To UBound(valuesD)
            ' Populate columns A and D in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 4).Value = Trim(valuesD(j))
        Next j
        
        ' Get values from column E and split by semicolon
        valuesE = Split(ws.Cells(i, 5).Value, ";")
        
        ' Loop through each value and populate columns A and E accordingly in "Result" sheet
        For j = 0 To UBound(valuesE)
            ' Populate columns A and E in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 5).Value = Trim(valuesE(j))
        Next j
        
        ' Get values from column F and split by semicolon
        valuesF = Split(ws.Cells(i, 6).Value, ";")
        
        ' Loop through each value and populate columns A and F accordingly in "Result" sheet
        For j = 0 To UBound(valuesF)
            ' Populate columns A and F in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 6).Value = Trim(valuesF(j))
        Next j
        
        ' Get values from column G and split by semicolon
        valuesG = Split(ws.Cells(i, 7).Value, ";")
        
        ' Loop through each value and populate columns A and G accordingly in "Result" sheet
        For j = 0 To UBound(valuesG)
            ' Populate columns A and G in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 7).Value = Trim(valuesG(j))
        Next j
        
        ' Get values from column H and split by semicolon
        valuesH = Split(ws.Cells(i, 8).Value, ";")
        
        ' Loop through each value and populate columns A and H accordingly in "Result" sheet
        For j = 0 To UBound(valuesH)
            ' Populate columns A and H in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 8).Value = Trim(valuesH(j))
        Next j
        
        ' Get values from column I and split by semicolon
        valuesI = Split(ws.Cells(i, 9).Value, ";")
        
        ' Loop through each value and populate columns A and I accordingly in "Result" sheet
        For j = 0 To UBound(valuesI)
            ' Populate columns A and I in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 9).Value = Trim(valuesI(j))
        Next j
        
        ' Get values from column J and split by semicolon
        valuesJ = Split(ws.Cells(i, 10).Value, ";")
        
        ' Loop through each value and populate columns A and J accordingly in "Result" sheet
        For j = 0 To UBound(valuesJ)
            ' Populate columns A and J in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 10).Value = Trim(valuesJ(j))
        Next j
        
        ' Get values from column K and split by semicolon
        valuesK = Split(ws.Cells(i, 11).Value, ";")
        
        ' Loop through each value and populate columns A and K accordingly in "Result" sheet
        For j = 0 To UBound(valuesK)
            ' Populate columns A and K in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 11).Value = Trim(valuesK(j))
        Next j
        
        ' Get values from column L and split by semicolon
        valuesL = Split(ws.Cells(i, 12).Value, ";")
        
        ' Loop through each value and populate columns A and L accordingly in "Result" sheet
        For j = 0 To UBound(valuesL)
            ' Populate columns A and L in the "Result" sheet
            resultWs.Cells(k + j, 1).Value = id
            resultWs.Cells(k + j, 12).Value = Trim(valuesL(j))
        Next j

        ' Increment the result row counter
        k = k + WorksheetFunction.Max(UBound(valuesB), UBound(valuesC), UBound(valuesD), UBound(valuesE), UBound(valuesF), UBound(valuesG), UBound(valuesH), UBound(valuesI), UBound(valuesJ), UBound(valuesK), UBound(valuesL)) + 1
    Next i
End Sub

