VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveIncomeExpense 
   Caption         =   "Remove Item From I&E"
   ClientHeight    =   5976
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4944
   OleObjectBlob   =   "RemoveIncomeExpense.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveIncomeExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub removeCategory_Change()
    ' Clear the removeItem combo box
    removeItem.Clear
    
    ' Populate removeItem based on the selected category
    If removeCategory.Value = "Income" Then
        removeItem.AddItem "Salary"
        removeItem.AddItem "Side Hustles"
        removeItem.AddItem "Bonus"
        removeItem.AddItem "Other"
    ElseIf removeCategory.Value = "Expense" Then
        removeItem.AddItem "Rent"
        removeItem.AddItem "Utilities"
        removeItem.AddItem "Food"
        removeItem.AddItem "Car"
        removeItem.AddItem "Gas"
        removeItem.AddItem "Bills"
        removeItem.AddItem "Shopping"
        removeItem.AddItem "Entertainment"
        removeItem.AddItem "Miscellaneous"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Populate the removeCategory combo box
    With removeCategory
        .AddItem "Income"
        .AddItem "Expense"
    End With
    
    ' Populate removeItem based on the default value
    Call removeCategory_Change
End Sub

Private Sub removeButton_Click()
    Dim ws As Worksheet
    Dim removeDate As Date
    Dim category As String, item As String
    Dim startRow As Long
    Dim lastRow As Long
    Dim currentRow As Long
    Dim targetRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Expenses&Incomes")

    ' Read user inputs from the UserForm
    With RemoveIncomeExpense
        removeDate = DateSerial(.removeYear.Text, .removeMonth.Text, .removeDay.Text)
        category = .removeCategory.Value
        item = .removeItem.Value
    End With

    ' Define the starting row for the table
    startRow = 2

    ' Find the last row in the table
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row

    ' ------------------------
    ' Clear Matching Data
    ' ------------------------
    For currentRow = startRow To lastRow
        If ws.Cells(currentRow, "A").Value = removeDate And _
           ws.Cells(currentRow, "B").Value = category And _
           ws.Cells(currentRow, "C").Value = item Then
            ws.Range("A" & currentRow & ":F" & currentRow).ClearContents
        End If
    Next currentRow

    ' ------------------------
    ' Shift Remaining Rows Up
    ' ------------------------
    targetRow = startRow
    For currentRow = startRow To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & currentRow & ":F" & currentRow)) > 0 Then
            If currentRow <> targetRow Then
                ws.Range("A" & currentRow & ":F" & currentRow).Copy ws.Range("A" & targetRow & ":F" & targetRow)
                ws.Range("A" & currentRow & ":F" & currentRow).ClearContents
            End If
            targetRow = targetRow + 1
        End If
    Next currentRow

    MsgBox "Matching data cleared and whitespace removed.", vbInformation
End Sub


Private Sub clearAllButton_Click()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Reference the worksheet
    Set ws = ThisWorkbook.Worksheets("Expenses&Incomes")
    
    ' Find the last row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Clear all data from columns A to F starting at row 2
    If lastRow >= 2 Then
        ws.Range("A2:F" & lastRow).ClearContents
    End If
    
    MsgBox "All data in columns A to F cleared.", vbInformation
End Sub


