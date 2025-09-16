Attribute VB_Name = "Module5"
Sub ExportAllShippingLabelsToPDF_V1()
    Dim wsProd As Worksheet, wsTemplate As Worksheet
    Dim fdcList As Range, cell As Range
    Dim fdcValue As String
    Dim tempSheet As Worksheet
    Dim exportPath As String
    Dim tempSheets As Collection
    Dim sheetName As String
    Dim i As Long
    Dim sheetArray() As Variant
    Dim t As Double
    Dim originalFormula As String

    Set wsProd = ThisWorkbook.Sheets("Production")
    Set wsTemplate = ThisWorkbook.Sheets("shipping label template")
    Set fdcList = wsProd.Range("AH5:AH" & wsProd.Cells(wsProd.Rows.Count, "AH").End(xlUp).Row)
    Set tempSheets = New Collection

    exportPath = ThisWorkbook.Path & "\All_Shipping_Labels.pdf"

    ' Backup original formula in F17
    originalFormula = wsTemplate.Range("F17").Formula

    For Each cell In fdcList
        fdcValue = Trim(cell.Value)
        If fdcValue <> "" Then
            wsTemplate.Range("A4").Value = fdcValue
            Application.CalculateFull
            Application.Wait Now + TimeValue("0:00:01")
            DoEvents

            t = Timer
            Do While wsTemplate.Range("D21").Value = "" And (Timer - t < 3)
                DoEvents
            Loop

            Call InsertSingleImageFromLabels_V1
            Application.Wait Now + TimeValue("0:00:01")
            DoEvents

            ' Always restore F17 formula before exporting
            wsTemplate.Range("F17").Formula = originalFormula

            wsTemplate.Copy After:=Sheets(Sheets.Count)
            Set tempSheet = ActiveSheet

            On Error Resume Next
            sheetName = "LABEL_" & fdcValue & "_single"
            tempSheet.Name = sheetName
            On Error GoTo 0

            tempSheets.Add tempSheet
        End If
    Next cell

    ' Export all sheets
    If tempSheets.Count > 0 Then
        ReDim sheetArray(1 To tempSheets.Count)
        For i = 1 To tempSheets.Count
            sheetArray(i) = tempSheets(i).Name
        Next i
        Sheets(sheetArray).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=exportPath, Quality:=xlQualityStandard
    End If

    Application.DisplayAlerts = False
    For Each tempSheet In tempSheets
        tempSheet.Delete
    Next tempSheet
    Application.DisplayAlerts = True

    MsgBox "All shipping labels saved to PDF:" & vbNewLine & exportPath, vbInformation
End Sub

