Attribute VB_Name = "Module4"
Sub GoToDashboard()
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Activate
    On Error GoTo 0
End Sub

Sub GoToExpensesAndIncome()
    On Error Resume Next
    ThisWorkbook.Sheets("Expenses&Incomes").Activate
    On Error GoTo 0
End Sub

Sub GoToOutput()
    On Error Resume Next
    ThisWorkbook.Sheets("Output").Activate
    On Error GoTo 0
End Sub

Sub GoToGoals()
    On Error Resume Next
    ThisWorkbook.Sheets("Goals").Activate
    On Error GoTo 0
End Sub

Sub GoToFinancialAdvice()
    On Error Resume Next
    ThisWorkbook.Sheets("Financial Advice").Activate
    On Error GoTo 0
End Sub
Sub GoToInstructions()
    On Error Resume Next
    ThisWorkbook.Sheets("Instructions").Activate
    On Error GoTo 0
End Sub

Sub RefreshPieChartsDash()
    Dim ws As Worksheet
    Dim pt As pivotTable

    ' Set the worksheet containing the pivot tables
    Set ws = ThisWorkbook.Sheets("Tracking Finances")

    ' Refresh specific pivot tables
    On Error Resume Next
    Set pt = ws.PivotTables("OutputPivotChartTF")
    If Not pt Is Nothing Then pt.RefreshTable

    Set pt = ws.PivotTables("IncomeAllocationPivotTable")
    If Not pt Is Nothing Then
        With pt
            ' Adjust the filter field name to match your PivotTable field
            .PivotFields("Category").CurrentPage = "Income"
            .RefreshTable
        End With
    End If

    ' Refresh and set filter for ExpenseAllocationPivotTable (Expense)
    Set pt = ws.PivotTables("ExpenseAllocationPivotTable")
    If Not pt Is Nothing Then
        With pt
            ' Adjust the filter field name to match your PivotTable field
            .PivotFields("Category").CurrentPage = "Expense"
            .RefreshTable
        End With
    End If
    On Error GoTo 0

    MsgBox "All specified pie charts have been refreshed!", vbInformation
End Sub

Sub RefreshDoubleBarGraphDash()
    Dim ws As Worksheet
    Dim pt As pivotTable

    ' Set the worksheet containing the pivot tables
    Set ws = ThisWorkbook.Sheets("Tracking Finances")

    ' Refresh specific pivot tables
    On Error Resume Next
    Set pt = ws.PivotTables("OutputPivotTableTF")
    If Not pt Is Nothing Then pt.RefreshTable
    On Error GoTo 0
    

    

    MsgBox "All specified charts have been refreshed!", vbInformation
End Sub

