Attribute VB_Name = "Module3"
Sub SendToTrackingFinances()
    Dim wsSource As Worksheet
    Dim wsTracking As Worksheet
    Dim lastRowSource As Long, lastRowIncome As Long, lastRowExpense As Long, lastRowCombined As Long
    Dim cell As Range
    Dim incomeRow As Long, expenseRow As Long, combinedRow As Long
    Dim recurrenceType As String
    Dim periodCount As Long
    Dim entryDate As Date
    Dim i As Long
    Dim category As String
    Dim similarDataFound As Boolean
    Dim promptUser As Boolean
    Dim userResponse As VbMsgBoxResult

    Set wsSource = ThisWorkbook.Sheets("Expenses&Incomes")
    Set wsTracking = ThisWorkbook.Sheets("Tracking Finances")
    
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).row
    
    lastRowIncome = wsTracking.Cells(wsTracking.Rows.Count, "A").End(xlUp).row
    If lastRowIncome < 3 Then lastRowIncome = 2
    incomeRow = lastRowIncome + 1
    
    lastRowExpense = wsTracking.Cells(wsTracking.Rows.Count, "F").End(xlUp).row
    If lastRowExpense < 3 Then lastRowExpense = 2
    expenseRow = lastRowExpense + 1
    
    lastRowCombined = wsTracking.Cells(wsTracking.Rows.Count, "K").End(xlUp).row
    If lastRowCombined < 3 Then lastRowCombined = 2
    combinedRow = lastRowCombined + 1
    
    For Each cell In wsSource.Range("D2:D" & lastRowSource)
        If IsNumeric(cell.Value) Then
            category = cell.Offset(0, -2).Value
            recurrenceType = cell.Offset(0, 1).Value
            periodCount = cell.Offset(0, 2).Value
            entryDate = wsSource.Cells(cell.row, 1).Value

            If Not IsDate(entryDate) Then
                MsgBox "Invalid date in row " & cell.row & ". Skipping entry.", vbExclamation
                GoTo NextIteration
            End If

            similarDataFound = False
            promptUser = False

            ' Check for similar data in the Tracking Finances sheet
            For Each tCell In wsTracking.Range("K3:K" & lastRowCombined)
                If tCell.Value = entryDate And _
                   tCell.Offset(0, 1).Value = category And _
                   tCell.Offset(0, 2).Value = cell.Offset(0, -1).Value And _
                   tCell.Offset(0, 3).Value = IIf(category = "Expense", cell.Value, cell.Value) Then
                   
                   similarDataFound = True
                   Exit For
                End If
            Next tCell

            ' Prompt user if similar data is found
            If similarDataFound Then
                userResponse = MsgBox("Similar data found in Tracking Finances for date: " & entryDate & _
                                      ", category: " & category & ", and item: " & cell.Offset(0, -1).Value & "." & vbCrLf & _
                                      "Do you still want to send this data?", vbYesNo + vbQuestion, "Confirm Sending Data")
                If userResponse = vbNo Then
                    GoTo NextIteration
                End If
            End If

            For i = 1 To IIf(periodCount > 0, periodCount, 1)
                If category = "Income" Then
                    wsTracking.Cells(incomeRow, 1).Value = entryDate
                    wsTracking.Cells(incomeRow, 2).Value = category
                    wsTracking.Cells(incomeRow, 3).Value = cell.Offset(0, -1).Value ' Item
                    wsTracking.Cells(incomeRow, 4).Value = cell.Value               ' Amount
                    incomeRow = incomeRow + 1
                ElseIf category = "Expense" Then
                    wsTracking.Cells(expenseRow, 6).Value = entryDate
                    wsTracking.Cells(expenseRow, 7).Value = category
                    wsTracking.Cells(expenseRow, 8).Value = cell.Offset(0, -1).Value ' Item
                    wsTracking.Cells(expenseRow, 9).Value = cell.Value               ' Amount
                    expenseRow = expenseRow + 1
                End If
                
                wsTracking.Cells(combinedRow, 11).Value = entryDate
                wsTracking.Cells(combinedRow, 12).Value = category
                wsTracking.Cells(combinedRow, 13).Value = cell.Offset(0, -1).Value   ' Item
                wsTracking.Cells(combinedRow, 14).Value = IIf(category = "Expense", cell.Value, cell.Value) ' Negative for Expense
                combinedRow = combinedRow + 1

                Select Case recurrenceType
                    Case "Daily"
                        entryDate = DateAdd("d", 1, entryDate)
                    Case "Weekly"
                        entryDate = DateAdd("d", 7, entryDate)
                    Case "Bi-Weekly"
                        entryDate = DateAdd("d", 14, entryDate)
                    Case "Monthly"
                        entryDate = DateAdd("m", 1, entryDate)
                    Case "Annually"
                        entryDate = DateAdd("yyyy", 1, entryDate)
                    Case Else
                        Exit For
                End Select
            Next i
        End If
NextIteration:
    Next cell
    
    wsTracking.Range("A1:F1").Font.Bold = True
    wsTracking.Range("G1:J1").Font.Bold = True
    wsTracking.Range("K1:N1").Font.Bold = True
    
    wsTracking.Range("A2:F2").EntireColumn.AutoFit
    wsTracking.Range("G2:J2").EntireColumn.AutoFit
    wsTracking.Range("K2:N2").EntireColumn.AutoFit

    With wsTracking
        .Range("A3:D" & incomeRow - 1).Sort Key1:=.Range("A3:A" & incomeRow - 1), Order1:=xlAscending, Header:=xlNo
        .Range("F3:I" & expenseRow - 1).Sort Key1:=.Range("F3:F" & expenseRow - 1), Order1:=xlAscending, Header:=xlNo
        .Range("K3:N" & combinedRow - 1).Sort Key1:=.Range("K3:K" & combinedRow - 1), Order1:=xlAscending, Header:=xlNo
    End With
End Sub


