Attribute VB_Name = "Module7"

Sub RefreshAllDashboardCharts()

    ' Call the macros to refresh the pie charts and double bar graph
    Call RefreshPieChartsDash
    Call RefreshDoubleBarGraphDash

    MsgBox "Dashboard charts have been refreshed successfully.", vbInformation, "Refresh Complete"

End Sub

