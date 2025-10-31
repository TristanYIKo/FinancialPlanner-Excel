Attribute VB_Name = "Module6"
Sub UpdateOutput()
    Dim WB As Workbook
    Dim ws As Worksheet
    Dim intReadRow As Integer
    Dim intWriteRow As Integer
    Dim intTrackingRow As Integer
    Dim startDate As Date
    Dim endDate As Date
    Dim recordDate As Date
    Dim trackingSheet As Worksheet
    Dim dashboardSheet As Worksheet
    Dim lastRow As Long
    Dim totalIncome As Double
    Dim totalExpenses As Double
    Dim netWorth As Double

    Set WB = ThisWorkbook
    Set ws = WB.Worksheets("Output")
    Set trackingSheet = WB.Worksheets("Tracking Finances")
    Set dashboardSheet = WB.Worksheets("Dashboard")

    intReadRow = 3
    intWriteRow = 2
    intTrackingRow = 3 ' Start writing on row 3 in Tracking Finances

    On Error GoTo DateError

    ' Read start and end dates from Dashboard text boxes
    startDate = CDate(dashboardSheet.Shapes("StartDateTextBox").TextFrame2.TextRange.Text)
    endDate = CDate(dashboardSheet.Shapes("EndDateTextBox").TextFrame2.TextRange.Text)
    On Error GoTo 0

    ' Place start date and end date in the Output sheet
    ws.Cells(2, "E").Value = startDate
    ws.Cells(4, "E").Value = endDate

    ' Clear existing data in the Output sheet (columns I to L)
    ws.Range("I2:L" & ws.Cells(ws.Rows.Count, "I").End(xlUp).row).ClearContents

    ' Clear existing data in the Tracking Finances sheet (columns AA to AD)
    lastRow = trackingSheet.Cells(trackingSheet.Rows.Count, "AA").End(xlUp).row
    If lastRow >= 3 Then
        trackingSheet.Range("AA3:AD" & lastRow).ClearContents
    End If

    ' Set headers for Output sheet
    ws.Cells(1, "I").Value = "Date"
    ws.Cells(1, "J").Value = "Category"
    ws.Cells(1, "K").Value = "Item"
    ws.Cells(1, "L").Value = "Amount"

    totalIncome = 0
    totalExpenses = 0

    ' Loop through Tracking Finances to filter data within the date range
    Do While trackingSheet.Cells(intReadRow, "A").Value <> "" Or trackingSheet.Cells(intReadRow, "F").Value <> ""

        ' Process income data
        If IsDate(trackingSheet.Cells(intReadRow, "A").Value) Then
            recordDate = trackingSheet.Cells(intReadRow, "A").Value
            If recordDate >= startDate And recordDate <= endDate Then
                ' Add to Output sheet
                ws.Cells(intWriteRow, "I").Value = recordDate
                ws.Cells(intWriteRow, "J").Value = trackingSheet.Cells(intReadRow, "B").Value
                ws.Cells(intWriteRow, "K").Value = trackingSheet.Cells(intReadRow, "C").Value
                ws.Cells(intWriteRow, "L").Value = trackingSheet.Cells(intReadRow, "D").Value
                intWriteRow = intWriteRow + 1

                ' Add to Tracking Finances sheet
                trackingSheet.Cells(intTrackingRow, "AA").Value = recordDate
                trackingSheet.Cells(intTrackingRow, "AB").Value = trackingSheet.Cells(intReadRow, "B").Value
                trackingSheet.Cells(intTrackingRow, "AC").Value = trackingSheet.Cells(intReadRow, "C").Value
                trackingSheet.Cells(intTrackingRow, "AD").Value = trackingSheet.Cells(intReadRow, "D").Value
                intTrackingRow = intTrackingRow + 1

                ' Accumulate total income
                totalIncome = totalIncome + trackingSheet.Cells(intReadRow, "D").Value
            End If
        End If

        ' Process expense data
        If IsDate(trackingSheet.Cells(intReadRow, "F").Value) Then
            recordDate = trackingSheet.Cells(intReadRow, "F").Value
            If recordDate >= startDate And recordDate <= endDate Then
                ' Add to Output sheet
                ws.Cells(intWriteRow, "I").Value = recordDate
                ws.Cells(intWriteRow, "J").Value = trackingSheet.Cells(intReadRow, "G").Value
                ws.Cells(intWriteRow, "K").Value = trackingSheet.Cells(intReadRow, "H").Value
                ws.Cells(intWriteRow, "L").Value = trackingSheet.Cells(intReadRow, "I").Value
                intWriteRow = intWriteRow + 1

                ' Add to Tracking Finances sheet
                trackingSheet.Cells(intTrackingRow, "AA").Value = recordDate
                trackingSheet.Cells(intTrackingRow, "AB").Value = trackingSheet.Cells(intReadRow, "G").Value
                trackingSheet.Cells(intTrackingRow, "AC").Value = trackingSheet.Cells(intReadRow, "H").Value
                trackingSheet.Cells(intTrackingRow, "AD").Value = trackingSheet.Cells(intReadRow, "I").Value
                intTrackingRow = intTrackingRow + 1

                ' Accumulate total expenses
                totalExpenses = totalExpenses + trackingSheet.Cells(intReadRow, "I").Value
            End If
        End If

        intReadRow = intReadRow + 1
    Loop

    ' Calculate net worth
    netWorth = totalIncome - totalExpenses

    ' Update the Dashboard net worth text box
    On Error Resume Next
    dashboardSheet.Shapes("netWorthText").TextFrame2.TextRange.Text = "$" & Format(netWorth, "0.00")
    On Error GoTo 0

    MsgBox "Data successfully updated and net worth refreshed.", vbInformation

    Exit Sub

DateError:
    MsgBox "Please enter valid dates in the Dashboard text boxes.", vbExclamation
End Sub

