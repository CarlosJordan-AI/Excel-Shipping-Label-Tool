Attribute VB_Name = "Module9"
Sub ExportBatchHeaderTemplates()
    Dim wsProd As Worksheet
    Dim wsGlobal As Worksheet, wsDepot As Worksheet, wsMulti As Worksheet
    Dim lastRow As Long, i As Long
    Dim outRowGlobal As Long, outRowDepot As Long, outRowMulti As Long
    Dim fdcDict As Object
    Dim fdcBU As String, qty As Variant

    Set wsProd = ThisWorkbook.Sheets("Production")
    Set wsGlobal = ThisWorkbook.Sheets("UPSGlobal")
    Set wsDepot = ThisWorkbook.Sheets("UPSHomeDepot")
    Set wsMulti = ThisWorkbook.Sheets("UPSMultiPackage")

    ' Clear previous output (preserve headers)
    wsGlobal.Range("A2:CB1000").ClearContents
    wsDepot.Range("A2:CB1000").ClearContents
    wsMulti.Range("A2:CB1000").ClearContents

    ' Use BU column as FDC key
    lastRow = wsProd.Cells(wsProd.Rows.Count, "BU").End(xlUp).Row
    outRowGlobal = 2: outRowDepot = 2: outRowMulti = 2

    ' Step 1: Count frequency of each FDC# in BU
    Set fdcDict = CreateObject("Scripting.Dictionary")
    For i = 5 To lastRow
        fdcBU = Trim(wsProd.Cells(i, "BU").Value)
        If fdcBU <> "" Then
            If Not fdcDict.exists(fdcBU) Then
                fdcDict(fdcBU) = 1
            Else
                fdcDict(fdcBU) = fdcDict(fdcBU) + 1
            End If
        End If
    Next i

    ' Step 2: Process and route rows based on BU
    For i = 5 To lastRow
        fdcBU = Trim(wsProd.Cells(i, "BU").Value)
        qty = wsProd.Cells(i, "P").Value

        If fdcBU <> "" Then
            Dim company As String: company = Trim(wsProd.Cells(i, "A").Value)
            Dim carrier As String: carrier = Trim(wsProd.Cells(i, "O").Value)
            Dim isMultipack As Boolean: isMultipack = False

            ' Multipackage rule: qty > 1 OR repeated FDC in BU
            If IsNumeric(qty) And qty > 1 Then
                isMultipack = True
            ElseIf fdcDict.exists(fdcBU) And fdcDict(fdcBU) > 1 Then
                isMultipack = True
            End If

            If carrier = "UPS" And (company = "GI" Or company = "CH") Then
                If isMultipack Then
                    WriteBatchHeaderFinal wsMulti, outRowMulti, wsProd, i, company, False
                    outRowMulti = outRowMulti + 1
                ElseIf company = "GI" Then
                    WriteBatchHeaderFinal wsGlobal, outRowGlobal, wsProd, i, company, True
                    outRowGlobal = outRowGlobal + 1
                ElseIf company = "CH" Then
                    WriteBatchHeaderFinal wsDepot, outRowDepot, wsProd, i, company, False
                    outRowDepot = outRowDepot + 1
                End If
            End If
        End If
    Next i

    ' Export populated sheets
    If outRowGlobal > 2 Then ExportSheetAsWorkbook wsGlobal, "UPSGlobal"
    If outRowDepot > 2 Then ExportSheetAsWorkbook wsDepot, "UPSHomeDepot"
    If outRowMulti > 2 Then ExportSheetAsWorkbook wsMulti, "UPSMultiPackage"

    MsgBox "UPS batch templates exported successfully.", vbInformation
End Sub

Private Sub WriteBatchHeaderFinal(ws As Worksheet, targetRow As Long, src As Worksheet, i As Long, company As String, isGlobal As Boolean)
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 80)).ClearContents

    With ws
        .Cells(targetRow, "A").Value = company                              ' Company
        .Cells(targetRow, "B").Value = src.Cells(i, "AN").Value            ' Company or Name
        .Cells(targetRow, "C").Value = "USA"                               ' Country
        .Cells(targetRow, "D").Value = src.Cells(i, "H").Value             ' Address 1
        .Cells(targetRow, "E").Value = src.Cells(i, "AO").Value            ' Address 2
        .Cells(targetRow, "G").Value = src.Cells(i, "K").Value             ' City
        .Cells(targetRow, "H").Value = src.Cells(i, "L").Value             ' State/Prov
        .Cells(targetRow, "I").Value = src.Cells(i, "I").Value             ' Postal Code
        .Cells(targetRow, "J").Value = src.Cells(i, "AR").Value            ' Telephone
        .Cells(targetRow, "L").Value = src.Cells(i, "BC").Value            ' Residential Ind
        .Cells(targetRow, "N").Value = 2                                   ' Packaging Type
        .Cells(targetRow, "P").Value = src.Cells(i, "AT").Value            ' Weight
        .Cells(targetRow, "Q").Value = src.Cells(i, "BD").Value            ' Length
        .Cells(targetRow, "R").Value = src.Cells(i, "BE").Value            ' Width
        .Cells(targetRow, "S").Value = src.Cells(i, "BF").Value            ' Height
        .Cells(targetRow, "U").Value = src.Cells(i, "BG").Value            ' Description of Goods
        .Cells(targetRow, "W").Value = 0                                   ' GNIFC
        .Cells(targetRow, "Y").Value = "'03"                               ' Service

        ' Reference logic
        If isGlobal Then
            .Cells(targetRow, "AG").Value = src.Cells(i, "B").Value        ' Ref1 from B
            .Cells(targetRow, "AH").Value = src.Cells(i, "C").Value        ' Ref2 from C
        ElseIf ws.Name = "UPSMultiPackage" And company = "GI" Then
            .Cells(targetRow, "AG").Value = src.Cells(i, "B").Value
            .Cells(targetRow, "AH").Value = src.Cells(i, "C").Value
        Else
            .Cells(targetRow, "AG").Value = src.Cells(i, "AW").Value
            .Cells(targetRow, "AH").Value = "'8119"
        End If
    End With
End Sub

Sub ExportSheetAsWorkbook(ws As Worksheet, fileName As String)
    Dim newWb As Workbook, newWs As Worksheet
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Sheets(1)

    ws.UsedRange.Copy Destination:=newWs.Range("A1")
    On Error Resume Next: newWs.Name = ws.Name: On Error GoTo 0

    newWb.SaveAs ThisWorkbook.Path & "\" & fileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

