Attribute VB_Name = "Module1"
Sub ClearExpensesIncomes()
    Dim wsSource As Worksheet
    Dim lastRow As Long
    
    Set wsSource = ThisWorkbook.Sheets("Expenses&Incomes")
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    
    If lastRow > 1 Then
        wsSource.Rows("2:" & lastRow).ClearContents
    End If
End Sub

Sub ClearOutput()

    ' Clear start date
    Range("E2").Select
    Selection.ClearContents
    
    ' Clear end date
    Range("E4").Select
    Selection.ClearContents
    
    ' Clear output data in the Output sheet
    Range("I2:L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    ' Clear output data in the Tracking Finances sheet
    Dim wsTracking As Worksheet
    Set wsTracking = ThisWorkbook.Worksheets("Tracking Finances")
    
    With wsTracking
        .Range("AA3:AD" & .Cells(.Rows.Count, "AA").End(xlUp).row).ClearContents
    End With

End Sub



Sub DeleteRowsFromTrackingFinances()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Tracking Finances")

    ' Find the last row with data in column K
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).row

    ' Loop through rows from the last to the first (starting at row 3)
    For currentRow = lastRow To 3 Step -1
        ws.Rows(currentRow).Delete
    Next currentRow

    MsgBox "All rows from K3 to N(last row) have been deleted.", vbInformation
End Sub

Sub ClearAllTables()
    Dim ws As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRowA As Long, lastRowF As Long, lastRowK As Long, lastRowOutput As Long
    Dim userResponse As VbMsgBoxResult

    ' Display a warning message before clearing the data
    userResponse = MsgBox("Are you sure you want to delete ALL data? This action cannot be undone.", _
                          vbYesNo + vbExclamation, "Warning: Delete All Data")

    ' If the user selects "No," exit the macro
    If userResponse = vbNo Then Exit Sub

    ' Reference the worksheets
    Set ws = ThisWorkbook.Sheets("Tracking Finances")
    Set outputSheet = ThisWorkbook.Sheets("Output")

    ' Find the last row for each table in Tracking Finances
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).row

    ' Clear the first table (Columns A to D)
    If lastRowA >= 3 Then
        ws.Range("A3:D" & lastRowA).ClearContents
    End If

    ' Clear the second table (Columns F to I)
    If lastRowF >= 3 Then
        ws.Range("F3:I" & lastRowF).ClearContents
    End If

    ' Clear the third table (Columns K to N)
    If lastRowK >= 3 Then
        ws.Range("K3:N" & lastRowK).ClearContents
    End If

    ' Clear the output table in the Output sheet (Columns I to L starting from row 2)
    lastRowOutput = outputSheet.Cells(outputSheet.Rows.Count, "I").End(xlUp).row
    If lastRowOutput >= 2 Then
        outputSheet.Range("I2:L" & lastRowOutput).ClearContents
    End If

    ' Clear the values in E2 and E4 on the Output sheet
    outputSheet.Range("E2").ClearContents
    outputSheet.Range("E4").ClearContents

    MsgBox "All data has been cleared.", vbInformation
End Sub


