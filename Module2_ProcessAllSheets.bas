Attribute VB_Name = "Module2"
' Module2 to apply processing on all sheets

' Steps:
' ----------------------------------------------------------------------------

Sub ProcessAllSheets()
    Dim ws As Worksheet
    ' Loop through each sheet in the active workbook
    For Each ws In ThisWorkbook.Sheets
        ' Perform an action, e.g., print the sheet name in the Immediate Window
        'MsgBox (ws.Name)
        Stock_per_Quarter_per_sheet ws
    Next ws
End Sub

