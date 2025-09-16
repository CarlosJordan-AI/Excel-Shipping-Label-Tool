Attribute VB_Name = "Module1"
Sub PreviewAndPrintSelectedFDC_V1()
    Dim wsProd As Worksheet
    Dim wsBOL As Worksheet
    Dim wsLabel As Worksheet
    Dim selectedFDC As String
    Dim printResponse As VbMsgBoxResult
    Dim printerChosen As Boolean

    Set wsProd = ThisWorkbook.Sheets("Production")
    Set wsBOL = ThisWorkbook.Sheets("bill of lading template")
    Set wsLabel = ThisWorkbook.Sheets("shipping label template")

    ' Get selected FDC# from dropdown
    On Error Resume Next
    selectedFDC = ThisWorkbook.Names("SelectedFDC").RefersToRange.Value
    On Error GoTo 0

    If selectedFDC = "" Then
        MsgBox "Please select an FDC# from the dropdown.", vbExclamation
        Exit Sub
    End If

    ' Set FDC# in AI2 to drive templates
    wsProd.Range("AI2").Value = selectedFDC

    ' Force formula update
    Application.CalculateFull
    DoEvents
    Application.Wait Now + TimeValue("0:00:01")

    ' Show print preview of both templates
    wsLabel.PrintPreview
    wsBOL.PrintPreview

    ' Ask if the user wants to print
    printResponse = MsgBox("Print these documents for FDC# " & selectedFDC & "?", vbYesNo + vbQuestion, "Confirm Print")
    If printResponse = vbNo Then Exit Sub

    ' Let user select printer
    printerChosen = Application.Dialogs(xlDialogPrinterSetup).Show
    If Not printerChosen Then Exit Sub

    ' Print both templates
    wsLabel.PrintOut Copies:=1
    wsBOL.PrintOut Copies:=1

    MsgBox "Printed documents for FDC#: " & selectedFDC, vbInformation
End Sub

