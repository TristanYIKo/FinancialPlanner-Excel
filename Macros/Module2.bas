Attribute VB_Name = "Module2"
Sub OpenOutputUserForm()
    OutputForm.Show
End Sub
Sub ShowAddItemUserForm()
    AddItemForm.Show
End Sub
Sub ShowRemoveItemUserForm()
    RemoveItemForm.Show
End Sub

Sub ShowAddMoneyToGoalsUserForm()
    AddMoneyToGoalsForm.Show
End Sub
Sub ShowRemoveIncomeExpenseUserForm()
    RemoveIncomeExpense.Show
End Sub
Sub ShowAdviceUserForm()
    FinancialAdviceForm.Show
End Sub

Sub FinancialAdvice_Button3_Click()

    Dim ws As Worksheet
    Dim incomeRange As Range
    Dim expenseRange As Range
    Dim totalIncome As Double
    Dim totalExpenses As Double
    Dim netWorth As Double
    Dim savingsGoal As Double
    Dim savingsPercentage As Double
    Dim message As String

    Set ws = ThisWorkbook.Sheets("Tracking Finances")
    Set incomeRange = ws.Range("D3:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).row)
    Set expenseRange = ws.Range("I3:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).row)

    totalIncome = Application.WorksheetFunction.Sum(incomeRange)
    totalExpenses = Application.WorksheetFunction.Sum(expenseRange)

    netWorth = totalIncome - totalExpenses
    savingsGoal = totalIncome * 0.2
    savingsPercentage = (totalIncome - totalExpenses) / totalIncome

    message = "Total Income: $" & Format(totalIncome, "0.00") & vbCrLf
    message = message & "Total Expenses: $" & Format(totalExpenses, "0.00") & vbCrLf
    message = message & "Net Worth (Income - Expenses): $" & Format(netWorth, "0.00") & vbCrLf
    message = message & "Recommended Savings (20% of Income): $" & Format(savingsGoal, "0.00") & vbCrLf

    If savingsPercentage >= 0.2 Then
        message = message & vbCrLf & "Congratulations! You are saving enough (at least 20% of your income)."
    Else
        message = message & vbCrLf & "You should aim to save more. Try to save at least 20% of your income."
    End If

    MsgBox message, vbInformation, "Savings and Financial Status"

End Sub

