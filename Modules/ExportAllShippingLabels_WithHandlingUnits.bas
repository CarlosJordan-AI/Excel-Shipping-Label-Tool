Attribute VB_Name = "Module7"
Public IsHandlingLabels_V1 As Boolean  ' Global flag

Sub ExportAllShippingLabels_WithHandlingUnits_V1()
    Dim wsProd As Worksheet, wsTemplate As Worksheet
    Dim fdcList As Range, cell As Range
    Dim fdcValue As String
    Dim tempSheet As Worksheet
    Dim exportPath As String
    Dim tempSheets As Collection
    Dim sheetName As String
    Dim i As Long, j As Long
    Dim sheetArray() As String
    Dim t As Double
    Dim handlingUnits As Long
    Dim huText As String
    Dim originalFormula As String

    Set wsProd = Sheets("Production")
    Set wsTemplate = Sheets("shipping label template")
    Set fdcList = wsProd.Range("AH5:AH" & wsProd.Cells(Rows.Count, "AH").End(xlUp).Row)
    Set tempSheets = New Collection

    exportPath = ThisWorkbook.Path & "\ShippingLabels_WithCopies.pdf"

    IsHandlingLabels_V1 = True  ' Activate handling-unit logic

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
            DoEvents

            originalFormula = wsTemplate.Range("F18").Formula
            huText = wsTemplate.Range("F18").Value
            handlingUnits = 1
            If InStr(huText, "of") > 0 Then handlingUnits = Val(Split(huText, "of")(1))

            For j = 1 To handlingUnits
                If handlingUnits > 1 Then
                    wsTemplate.Range("F18").Value = j & " of " & handlingUnits
                End If

                wsTemplate.Copy After:=Sheets(Sheets.Count)
                Set tempSheet = ActiveSheet

                sheetName = "LABEL_" & fdcValue & "_unit" & j
                On Error Resume Next
                tempSheet.Name = sheetName
                On Error GoTo 0

                tempSheets.Add tempSheet
            Next j

            wsTemplate.Range("F18").Formula = originalFormula
        End If
    Next cell

    IsHandlingLabels_V1 = False  ' Reset flag

    If tempSheets.Count > 0 Then
        ReDim sheetArray(1 To tempSheets.Count)
        For i = 1 To tempSheets.Count
            sheetArray(i) = tempSheets(i).Name
        Next i
        Sheets(sheetArray).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        fileName:=exportPath, _
                                        Quality:=xlQualityStandard
    End If

    Application.DisplayAlerts = False
    For Each tempSheet In tempSheets
        tempSheet.Delete
    Next tempSheet
    Application.DisplayAlerts = True

    MsgBox "Shipping labels with handling units saved to:" & vbNewLine & exportPath, _
           vbInformation
End Sub


