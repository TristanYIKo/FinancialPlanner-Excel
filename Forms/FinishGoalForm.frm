VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FinishGoalForm 
   Caption         =   "Complete Goal"
   ClientHeight    =   2648
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   6856
   OleObjectBlob   =   "FinishGoalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FinishGoalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitBtnGoalsFinish_Click()

    Dim WB As Workbook
    Dim ws As Worksheet
    Dim goalName As String
    Dim lastRow As Long
    Dim i As Long
    Dim goalFound As Boolean
    
    ' Set the workbook and worksheet
    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Goals")

    ' Get the goal name from the user form
    goalName = Trim(finishedGoalNameBox.Value)

    ' Validate the input
    If goalName = "" Then
        MsgBox "Please enter a goal name.", vbExclamation
        Exit Sub
    End If
    
    ' Find the last row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    goalFound = False
    
    ' Loop through rows to find the matching goal
    For i = 2 To lastRow
        If ws.Cells(i, "C").Value = goalName Then
            ' Clear columns C to I for the matching row
            ws.Range("C" & i & ":I" & i).ClearContents
            
            ' Shift rows below up
            Dim j As Long
            For j = i To lastRow - 1
                ws.Range("C" & j & ":I" & j).Value = ws.Range("C" & j + 1 & ":I" & j + 1).Value
            Next j
            
            ' Clear the last row after shifting
            ws.Range("C" & lastRow & ":I" & lastRow).ClearContents
            
            goalFound = True
            Exit For
        End If
    Next i
    
    ' Provide feedback to the user
    If goalFound Then
        MsgBox "Congratulations on finishing your goal!", vbInformation
    Else
        MsgBox "Goal not found. Please check the goal name and try again.", vbExclamation
    End If

End Sub

