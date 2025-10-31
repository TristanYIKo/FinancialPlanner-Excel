VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinancialAdviceForm 
   Caption         =   "Finacial Advice"
   ClientHeight    =   4592
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4744
   OleObjectBlob   =   "FinancialAdviceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinancialAdviceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' Add options to the adviceBox combo box
    With adviceBox
        .Clear
        .AddItem "income"
        .AddItem "spending"
        .AddItem "general"
    End With

    ' Add options to the adviceDate combo box
    With adviceDate
        .Clear
        .AddItem "all time"
        .AddItem "output range"
    End With
End Sub

Private Sub adviceButton_Click()

    Dim ws As Worksheet
    Dim outputSheet As Worksheet
    Dim incomeRange As Range
    Dim expenseRange As Range
    Dim totalIncome As Double
    Dim totalExpenses As Double
    Dim netSavings As Double
    Dim percentage As Double
    Dim adviceOption As String
    Dim dateOption As String
    Dim startDate As Date
    Dim endDate As Date
    Dim row As Range
    Dim message As String
    Dim randomAdvice As String

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Tracking Finances")
    Set outputSheet = ThisWorkbook.Sheets("Output")

    ' Get the selected options
    adviceOption = adviceBox.Value
    dateOption = adviceDate.Value

    ' Initialize totals
    totalIncome = 0
    totalExpenses = 0

    ' Handle date filtering
    If dateOption = "output range" Then
        On Error Resume Next ' Prevent crash if invalid date is entered
        startDate = CDate(outputSheet.Range("E2").Value)
        endDate = CDate(outputSheet.Range("E4").Value)
        On Error GoTo 0

        If startDate = 0 Or endDate = 0 Then
            MsgBox "Please enter valid start and end dates in the Output sheet.", vbExclamation, "Invalid Date Range"
            Exit Sub
        End If

        If startDate > endDate Then
            MsgBox "Start date cannot be after the end date.", vbExclamation, "Invalid Date Range"
            Exit Sub
        End If

        ' Loop through the income and expense ranges and sum only those within the date range
        For Each row In ws.Range("A3:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).row)
            If IsDate(row.Value) Then
                If row.Value >= startDate And row.Value <= endDate Then
                    totalIncome = totalIncome + ws.Cells(row.row, "D").Value
                    totalExpenses = totalExpenses + ws.Cells(row.row, "I").Value
                End If
            End If
        Next row
    Else
        ' Sum all data for "all time"
        totalIncome = Application.WorksheetFunction.Sum(ws.Range("D3:D" & ws.Cells(ws.Rows.Count, "D").End(xlUp).row))
        totalExpenses = Application.WorksheetFunction.Sum(ws.Range("I3:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).row))
    End If

    ' Calculate net savings and percentage
    netSavings = totalIncome - totalExpenses
    If totalIncome > 0 Then
        percentage = (netSavings / totalIncome) * 100
    Else
        percentage = 0
    End If

    ' Provide advice based on the selected option
    Select Case adviceOption
        Case "income"
            If totalIncome < totalExpenses Then
                message = "Your expenses exceed your income. You should focus on saving more money and reducing unnecessary expenses." & vbCrLf
            Else
                message = "Your income exceeds your expenses. Keep saving and building your financial security." & vbCrLf
            End If
            message = message & "Income: $" & Format(totalIncome, "0.00") & vbCrLf
            message = message & "Expenses: $" & Format(totalExpenses, "0.00") & vbCrLf
            message = message & "Net Savings: $" & Format(netSavings, "0.00") & " (" & Format(percentage, "0.00") & "%)"

        Case "spending"
            If totalExpenses > totalIncome Then
                message = "Your spending is more than your income. Stop spending excessively and reevaluate your budget." & vbCrLf
            Else
                message = "Your spending is less than your income. Consider investing the surplus to grow your wealth." & vbCrLf
            End If
            message = message & "Income: $" & Format(totalIncome, "0.00") & vbCrLf
            message = message & "Expenses: $" & Format(totalExpenses, "0.00") & vbCrLf
            message = message & "Net Savings: $" & Format(netSavings, "0.00") & " (" & Format(percentage, "0.00") & "%)"

        Case "general"
            ' Random financial advice
            Randomize
            Select Case Int((5 - 1 + 1) * Rnd + 1) ' Generate a random number between 1 and 5
                Case 1: randomAdvice = "Start an emergency fund with at least 3-6 months of living expenses."
                Case 2: randomAdvice = "Track your spending to identify areas where you can cut costs."
                Case 3: randomAdvice = "Invest in low-cost index funds to grow your wealth over time."
                Case 4: randomAdvice = "Avoid high-interest debt and pay off credit cards in full each month."
                Case 5: randomAdvice = "Review your financial goals regularly to stay on track."
            End Select
            message = randomAdvice

        Case Else
            message = "Please select a valid option from the advice box."
    End Select

    ' Display the message
    MsgBox message, vbInformation, "Financial Advice"

End Sub

