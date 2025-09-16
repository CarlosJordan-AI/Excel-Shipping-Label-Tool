Attribute VB_Name = "Module4"
Sub PreviewAndPrintShippingLabel_V1()
    Dim wsLabel As Worksheet
    Dim selectedFDC As String
    Dim printerChosen As Boolean
    Dim printResponse As VbMsgBoxResult

    Set wsLabel = ThisWorkbook.Sheets("shipping label template")
    selectedFDC = wsLabel.Range("A4").Value

    If selectedFDC = "" Then
        MsgBox "No FDC# is selected in the Shipping Label template (A4).", vbExclamation
        Exit Sub
    End If

    ' Preview
    wsLabel.PrintPreview

    ' Ask to print
    printResponse = MsgBox("Print Shipping Label for FDC# " & selectedFDC & "?", vbYesNo + vbQuestion)
    If printResponse = vbNo Then Exit Sub

    ' Let user choose printer
    printerChosen = Application.Dialogs(xlDialogPrinterSetup).Show
    If Not printerChosen Then Exit Sub

    ' Print
    wsLabel.PrintOut Copies:=1

    MsgBox "Shipping Label printed for FDC#: " & selectedFDC, vbInformation
End Sub

Sub PreviewAndPrintBillOfLading_V1()
    Dim wsBOL As Worksheet
    Dim selectedFDC As String
    Dim printerChosen As Boolean
    Dim printResponse As VbMsgBoxResult

    Set wsBOL = ThisWorkbook.Sheets("bill of lading template")
    selectedFDC = wsBOL.Range("AI2").Value

    If selectedFDC = "" Then
        MsgBox "No FDC# is selected in the Bill of Lading template (AI2).", vbExclamation
        Exit Sub
    End If

    ' Preview
    wsBOL.PrintPreview

    ' Ask to print
    printResponse = MsgBox("Print Bill of Lading for FDC# " & selectedFDC & "?", vbYesNo + vbQuestion)
    If printResponse = vbNo Then Exit Sub

    ' Let user choose printer
    printerChosen = Application.Dialogs(xlDialogPrinterSetup).Show
    If Not printerChosen Then Exit Sub

    ' Print
    wsBOL.PrintOut Copies:=1

    MsgBox "Bill of Lading printed for FDC#: " & selectedFDC, vbInformation
End Sub

