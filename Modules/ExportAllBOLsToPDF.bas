Attribute VB_Name = "Module8"
Sub ExportAllBOLsToPDF_V1()
    Dim wsProd As Worksheet, wsTemplate As Worksheet
    Dim fdcList As Range, cell As Range
    Dim fdcValue As String, tempSheet As Worksheet
    Dim exportPath As String, tempSheets As Collection
    Dim sheetName As String, i As Long
    Dim sheetArray() As Variant, t As Double

    Set wsProd = ThisWorkbook.Sheets("Production")
    Set wsTemplate = ThisWorkbook.Sheets("bill of lading template")
    Set fdcList = wsProd.Range("AH5:AH" & wsProd.Cells(wsProd.Rows.Count, "AH").End(xlUp).Row)
    Set tempSheets = New Collection

    exportPath = ThisWorkbook.Path & "\All_Bill_of_Lading.pdf"

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each cell In fdcList
        fdcValue = Trim(cell.Value)
        If fdcValue <> "" Then
            ' Directly assign FDC# to AI2, skipping AN1
            wsTemplate.Range("AI2").Formula = ""       ' Clear formula
            wsTemplate.Range("AI2").Value = fdcValue   ' Set directly

            ' Force recalc and wait for feed to update
            Application.CalculateFull
            Application.Wait Now + TimeValue("0:00:01")
            DoEvents

            ' Wait for one populated cell from feed (e.g., AJ3)
            t = Timer
            Do While wsTemplate.Range("AJ3").Value = "" And (Timer - t < 3)
                DoEvents
            Loop

            ' Copy sheet
            wsTemplate.Copy After:=Sheets(Sheets.Count)
            Set tempSheet = ActiveSheet

            On Error Resume Next
            sheetName = "BOL_" & fdcValue
            tempSheet.Name = sheetName
            On Error GoTo 0

            tempSheets.Add tempSheet
        End If
    Next cell

    ' Export all to PDF
    If tempSheets.Count > 0 Then
        ReDim sheetArray(1 To tempSheets.Count)
        For i = 1 To tempSheets.Count
            sheetArray(i) = tempSheets(i).Name
        Next i
        Sheets(sheetArray).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=exportPath, Quality:=xlQualityStandard
    End If

    ' Clean up
    Application.DisplayAlerts = False
    For Each tempSheet In tempSheets
        tempSheet.Delete
    Next tempSheet
    Application.DisplayAlerts = True

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "All BOLs saved to PDF:" & vbNewLine & exportPath, vbInformation
End Sub

