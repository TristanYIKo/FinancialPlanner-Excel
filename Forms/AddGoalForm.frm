VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddGoalForm 
   Caption         =   "Add Goals"
   ClientHeight    =   6736
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   10200
   OleObjectBlob   =   "AddGoalForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddGoalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    goalTypeBox.Clear
    goalTypeBox.AddItem "Save"
End Sub

Private Sub SubmitBtnGoals_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim goalName As String
    Dim goalType As String
    Dim goalDate As Date
    Dim goalAmount As Double
    Dim i As Long
    Dim goalCount As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Goals")

    ' Get data from the UserForm
    With AddGoalForm
        goalName = Trim(.goalNameBox.Value)
        goalType = Trim(.goalTypeBox.Value)
        
        ' Validate and construct the date
        On Error GoTo InvalidDate
        goalDate = DateSerial(CInt(.txtYear.Value), CInt(.txtMonth.Value), CInt(.txtDay.Value))
        On Error GoTo 0
        
        ' Validate the amount
        If Not IsNumeric(.amountBox.Value) Or Val(.amountBox.Value) <= 0 Then
            MsgBox "Please enter a valid positive amount.", vbExclamation
            Exit Sub
        End If
        goalAmount = CDbl(.amountBox.Value)
    End With

    ' Validate inputs
    If goalName = "" Then
        MsgBox "Please enter a goal name.", vbExclamation
        Exit Sub
    End If
    If goalType = "" Then
        MsgBox "Please select a goal type.", vbExclamation
        Exit Sub
    End If

    ' Find the last row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row
    
    ' Count existing goals
    goalCount = Application.WorksheetFunction.CountA(ws.Range("C2:C" & lastRow))
    If goalCount >= 4 Then
        MsgBox "You can only have a maximum of 4 goals.", vbExclamation
        Exit Sub
    End If

    ' Check for duplicate goal names
    For i = 2 To lastRow
        If LCase(ws.Cells(i, "C").Value) = LCase(goalName) Then
            MsgBox "A goal with this name already exists. Please choose a different name.", vbExclamation
            Exit Sub
        End If
    Next i

    ' Add data to the next available row
    ws.Cells(lastRow + 1, "C").Value = goalName
    ws.Cells(lastRow + 1, "D").Value = goalType
    ws.Cells(lastRow + 1, "E").Value = goalDate
    ws.Cells(lastRow + 1, "E").NumberFormat = "mmmm d, yyyy" ' Format date as "Month Day, Year"
    ws.Cells(lastRow + 1, "F").Value = goalAmount
    ws.Cells(lastRow + 1, "G").Value = goalAmount
    ws.Cells(lastRow + 1, "H").Value = 0
    ws.Cells(lastRow + 1, "H").NumberFormat = "0.00%"
    ws.Cells(lastRow + 1, "I").Value = 1
    ws.Cells(lastRow + 1, "I").NumberFormat = "0.00%"

    MsgBox "Goal successfully added!", vbInformation
    Exit Sub

InvalidDate:
    MsgBox "Please enter a valid date.", vbExclamation
End Sub

