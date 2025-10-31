VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveItemForm 
   Caption         =   "Remove Item"
   ClientHeight    =   5992
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   9296.001
   OleObjectBlob   =   "RemoveItemForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboxCategoryRemove_Change()
    RemoveItemForm.txtItemRemove.Clear
    If cboxCategoryRemove.Value = "Income" Then
        RemoveItemForm.txtItemRemove.AddItem "Salary"
        RemoveItemForm.txtItemRemove.AddItem "Side Hustles"
        RemoveItemForm.txtItemRemove.AddItem "Bonus"
        RemoveItemForm.txtItemRemove.AddItem "Other"
    ElseIf cboxCategoryRemove.Value = "Expense" Then
        RemoveItemForm.txtItemRemove.AddItem "Rent"
        RemoveItemForm.txtItemRemove.AddItem "Utilities"
        RemoveItemForm.txtItemRemove.AddItem "Food"
        RemoveItemForm.txtItemRemove.AddItem "Car"
        RemoveItemForm.txtItemRemove.AddItem "Gas"
        RemoveItemForm.txtItemRemove.AddItem "Bills"
        RemoveItemForm.txtItemRemove.AddItem "Shopping"
        RemoveItemForm.txtItemRemove.AddItem "Entertainment"
        RemoveItemForm.txtItemRemove.AddItem "Miscellaneous"
    End If
End Sub



Private Sub UserForm_Initialize()
    With RemoveItemForm.cboxCategoryRemove
        .AddItem "Income"
        .AddItem "Expense"
    End With
    
    Call cboxCategoryRemove_Change
End Sub

Private Sub SubmitBtnRemove_Click()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date
    Dim category As String, item As String
    Dim combinedStartRow As Long
    Dim lastRow As Long
    Dim currentRow As Long
    Dim targetRow As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Tracking Finances")

    ' Read user inputs from the UserForm
    With RemoveItemForm
        startDate = DateSerial(.txtYear1Remove.Value, .txtMonth1Remove.Value, .txtDay1Remove.Value)
        endDate = DateSerial(.txtYear2Remove.Value, .txtMonth2Remove.Value, .txtDay2Remove.Value)
        category = .cboxCategoryRemove.Value
        item = .txtItemRemove.Value
    End With

    ' Define starting row for the tables
    combinedStartRow = 3

    ' ------------------------
    ' Process Table in A:D
    ' ------------------------
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    For currentRow = combinedStartRow To lastRow
        If ws.Cells(currentRow, "A").Value >= startDate And ws.Cells(currentRow, "A").Value <= endDate Then
            If ws.Cells(currentRow, "B").Value = category And ws.Cells(currentRow, "C").Value = item Then
                ws.Range("A" & currentRow & ":D" & currentRow).ClearContents
            End If
        End If
    Next currentRow

    targetRow = combinedStartRow
    For currentRow = combinedStartRow To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("A" & currentRow & ":D" & currentRow)) > 0 Then
            If currentRow <> targetRow Then
                ws.Range("A" & currentRow & ":D" & currentRow).Copy ws.Range("A" & targetRow & ":D" & targetRow)
                ws.Range("A" & currentRow & ":D" & currentRow).ClearContents
            End If
            targetRow = targetRow + 1
        End If
    Next currentRow

    ' ------------------------
    ' Process Table in F:I
    ' ------------------------
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
    For currentRow = combinedStartRow To lastRow
        If ws.Cells(currentRow, "F").Value >= startDate And ws.Cells(currentRow, "F").Value <= endDate Then
            If ws.Cells(currentRow, "G").Value = category And ws.Cells(currentRow, "H").Value = item Then
                ws.Range("F" & currentRow & ":I" & currentRow).ClearContents
            End If
        End If
    Next currentRow

    targetRow = combinedStartRow
    For currentRow = combinedStartRow To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("F" & currentRow & ":I" & currentRow)) > 0 Then
            If currentRow <> targetRow Then
                ws.Range("F" & currentRow & ":I" & currentRow).Copy ws.Range("F" & targetRow & ":I" & targetRow)
                ws.Range("F" & currentRow & ":I" & currentRow).ClearContents
            End If
            targetRow = targetRow + 1
        End If
    Next currentRow

    ' ------------------------
    ' Process Table in K:N
    ' ------------------------
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).row
    For currentRow = combinedStartRow To lastRow
        If ws.Cells(currentRow, "K").Value >= startDate And ws.Cells(currentRow, "K").Value <= endDate Then
            If ws.Cells(currentRow, "L").Value = category And ws.Cells(currentRow, "M").Value = item Then
                ws.Range("K" & currentRow & ":N" & currentRow).ClearContents
            End If
        End If
    Next currentRow

    targetRow = combinedStartRow
    For currentRow = combinedStartRow To lastRow
        If Application.WorksheetFunction.CountA(ws.Range("K" & currentRow & ":N" & currentRow)) > 0 Then
            If currentRow <> targetRow Then
                ws.Range("K" & currentRow & ":N" & currentRow).Copy ws.Range("K" & targetRow & ":N" & targetRow)
                ws.Range("K" & currentRow & ":N" & currentRow).ClearContents
            End If
            targetRow = targetRow + 1
        End If
    Next currentRow

    MsgBox "Matching data cleared and whitespace removed from all tables.", vbInformation
End Sub

