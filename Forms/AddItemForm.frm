VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddItemForm 
   Caption         =   "Add Expense/Income"
   ClientHeight    =   5672
   ClientLeft      =   120
   ClientTop       =   464
   ClientWidth     =   10000
   OleObjectBlob   =   "AddItemForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboxCategory_Change()
    AddItemForm.txtItem.Clear
    If cboxCategory.Value = "Income" Then
        AddItemForm.txtItem.AddItem "Salary"
        AddItemForm.txtItem.AddItem "Side Hustles"
        AddItemForm.txtItem.AddItem "Bonus"
        AddItemForm.txtItem.AddItem "Other"
    ElseIf cboxCategory.Value = "Expense" Then
        AddItemForm.txtItem.AddItem "Rent"
        AddItemForm.txtItem.AddItem "Utilities"
        AddItemForm.txtItem.AddItem "Food"
        AddItemForm.txtItem.AddItem "Car"
        AddItemForm.txtItem.AddItem "Gas"
        AddItemForm.txtItem.AddItem "Bills"
        AddItemForm.txtItem.AddItem "Shopping"
        AddItemForm.txtItem.AddItem "Entertainment"
        AddItemForm.txtItem.AddItem "Miscellaneous"
    End If
End Sub



Private Sub UserForm_Initialize()
    With AddItemForm.cboxCategory
        .AddItem "Income"
        .AddItem "Expense"
    End With
    Call cboxCategory_Change
    With AddItemForm.recurringBox
        .AddItem "Daily"
        .AddItem "Weekly"
        .AddItem "Bi-Weekly"
        .AddItem "Monthly"
        .AddItem "Annually"
    End With
End Sub

Private Sub SubmitBtn_Click()
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Expenses&Incomes")
    intRow = 2
    If (txtItem.Value <> "") Then
        If (IsNumeric(txtDay.Value) And IsNumeric(txtMonth.Value) And IsNumeric(txtYear.Value)) Then
            If IsDate(txtYear.Value & "-" & txtMonth.Value & "-" & txtDay.Value) Then
                If (cboxCategory.Value <> "") Then
                    If (txtItem.Value <> "") Then
                        If txtDescription.Value = "" Or Not IsNumeric(txtDescription.Value) Then
                            MsgBox "Please enter a valid numeric amount.", vbExclamation
                            Exit Sub
                        End If
                        If recurringBox.Value <> "" Then
                            If periodBox.Value = "" Or Not IsNumeric(periodBox.Value) Then
                                MsgBox "Please enter a valid numeric period for the selected recurrence.", vbExclamation
                                Exit Sub
                            End If
                        End If
                        Do While (ws.Cells(intRow, "A") <> "")
                            intRow = intRow + 1
                        Loop
                        ws.Cells(intRow, "A") = txtYear.Value & "-" & txtMonth.Value & "-" & txtDay.Value
                        ws.Cells(intRow, "A").NumberFormat = "yyyy-mm-dd;@"
                        ws.Cells(intRow, "B") = cboxCategory.Value
                        ws.Cells(intRow, "C") = txtItem.Value
                        ws.Cells(intRow, "D") = txtDescription.Value
                        ws.Cells(intRow, "D").NumberFormat = "$#,##0.00"
                        If (recurringBox.Value <> "") Then
                            ws.Cells(intRow, "E") = recurringBox.Value
                        End If
                        ws.Cells(intRow, "F") = periodBox.Value
                    Else
                        MsgBox "Please select an item", vbExclamation
                    End If
                Else
                    MsgBox "Please select a category", vbExclamation
                End If
            Else
                MsgBox "Please enter a valid date.", vbExclamation
            End If
        Else
            MsgBox "Please ensure day, month, and year are numeric values.", vbExclamation
        End If
    Else
        MsgBox "Please enter an item", vbExclamation
    End If
End Sub

