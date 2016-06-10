
Private Sub Workbook_Activate()
    On Error Resume Next
    ActiveSheet.Worksheet_Activate
    On Error GoTo 0
End Sub

Private Sub Workbook_Deactivate()
    Application.StatusBar = ""
End Sub

