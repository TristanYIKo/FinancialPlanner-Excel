VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddMoneyToGoalsForm 
   Caption         =   "Goals Progress"
   ClientHeight    =   4624
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4112
   OleObjectBlob   =   "AddMoneyToGoalsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddMoneyToGoalsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub addMoneyGoalsBtn_Click()
    Dim ws As Worksheet
    Dim dashboard As Worksheet
    Dim goalName As String
    Dim amountToDeduct As Double
    Dim lastRow As Long
    Dim i As Long
    Dim goalFound As Boolean
    Dim currentNetWorth As Double

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Goals")
    Set dashboard = ThisWorkbook.Sheets("Dashboard")

    ' Get data from the UserForm
    With AddMoneyToGoalsForm
        goalName = Trim(.addMoneyGoalName.Value)
        If Not IsNumeric(.addMoneyGoalAmount.Value) Or Val(.addMoneyGoalAmount.Value) <= 0 Then
            MsgBox "Please enter a valid positive amount.", vbExclamation
            Exit Sub
        End If
        amountToDeduct = CDbl(.addMoneyGoalAmount.Value)
    End With

    ' Check if goal name is provided
    If goalName = "" Then
        MsgBox "Please enter a goal name.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    goalFound = False

    ' Search for the goal name in column C
    For i = 2 To lastRow
        If LCase(ws.Cells(i, "C").Value) = LCase(goalName) Then
            ' If goal name matches, validate and deduct the amount from column G
            If IsNumeric(ws.Cells(i, "G").Value) Then
                If amountToDeduct > ws.Cells(i, "G").Value Then
                    MsgBox "The amount entered exceeds the remaining money for this goal.", vbExclamation
                    Exit Sub
                End If

                ' Deduct the amount
                ws.Cells(i, "G").Value = ws.Cells(i, "G").Value - amountToDeduct

                ' If the goal balance reaches 0, erase the goal and move rows up
                If ws.Cells(i, "G").Value = 0 Then
                    Dim j As Long
                    For j = i To lastRow - 1
                        ws.Range("C" & j & ":I" & j).Value = ws.Range("C" & j + 1 & ":I" & j + 1).Value
                    Next j
                    ' Clear the last row after shifting
                    ws.Range("C" & lastRow & ":I" & lastRow).ClearContents
                    MsgBox "Goal has been completed and removed.", vbInformation
                Else
                    MsgBox "Amount successfully deducted from the goal.", vbInformation
                End If

                ' Update H and I columns after all changes to G
                If IsNumeric(ws.Cells(i, "F").Value) And ws.Cells(i, "F").Value > 0 Then
                    ws.Cells(i, "H").Value = (ws.Cells(i, "F").Value - ws.Cells(i, "G").Value) / ws.Cells(i, "F").Value
                    ws.Cells(i, "H").NumberFormat = "0.00%"
                    ws.Cells(i, "I").Value = ws.Cells(i, "G").Value / ws.Cells(i, "F").Value
                    ws.Cells(i, "I").NumberFormat = "0.00%"
                Else
                    ws.Cells(i, "H").Value = ""
                    ws.Cells(i, "I").Value = ""
                End If

                ' Update net worth in the Dashboard
                On Error Resume Next
                If IsNumeric(dashboard.Shapes("netWorthText").TextFrame2.TextRange.Text) Then
                    currentNetWorth = CDbl(Replace(dashboard.Shapes("netWorthText").TextFrame2.TextRange.Text, "$", ""))
                    currentNetWorth = currentNetWorth - amountToDeduct
                    dashboard.Shapes("netWorthText").TextFrame2.TextRange.Text = "$" & Format(currentNetWorth, "0.00")
                Else
                    MsgBox "Net Worth text box on Dashboard contains invalid data.", vbExclamation
                End If
                On Error GoTo 0

                goalFound = True
                Exit For
            Else
                MsgBox "The target cell in column G is not a valid number.", vbExclamation
                Exit Sub
            End If
        End If
    Next i

    ' If no matching goal name is found
    If Not goalFound Then
        MsgBox "Goal not found. Please check the goal name and try again.", vbExclamation
    End If

End Sub


