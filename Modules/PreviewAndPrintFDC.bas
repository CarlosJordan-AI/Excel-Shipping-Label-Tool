Attribute VB_Name = "Module2"
Sub PreviewAndPrintFDC_V1()
    Dim wsBOL As Worksheet
    Dim wsLabel As Worksheet
    Dim selectedFDC As String
    Dim printResponse As VbMsgBoxResult
    Dim printerChosen As Boolean

    Set wsBOL = ThisWorkbook.Sheets("bill of lading template")
    Set wsLabel = ThisWorkbook.Sheets("shipping label template")

    ' Get current FDC# from either template (both are synced)
    selectedFDC = wsBOL.Range("AI2").Value
    If selectedFDC = "" Then
        MsgBox "No FDC# is selected.", vbExclamation
        Exit Sub
    End If

    ' Show previews
    wsLabel.PrintPreview
    wsBOL.PrintPreview

    ' Confirm print
    printResponse = MsgBox("Print documents for FDC# " & selectedFDC & "?", vbYesNo + vbQuestion, "Confirm Print")
    If printResponse = vbNo Then Exit Sub

    ' Let user pick printer
    printerChosen = Application.Dialogs(xlDialogPrinterSetup).Show
    If Not printerChosen Then Exit Sub

    ' Print both
    wsLabel.PrintOut Copies:=1
    wsBOL.PrintOut Copies:=1

    MsgBox "Printed FDC#: " & selectedFDC, vbInformation
End Sub

