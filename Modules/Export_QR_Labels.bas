Attribute VB_Name = "Module10"
Sub Export_QR_Labels()
    Const SLOT_COUNT As Long = 60
    Const START_ROW As Long = 5
    Const TEMPLATE_SHEET As String = "QRs Labels"
    Const OUTPUT_NAME As String = "QR_Labels.pdf"

    Dim wsProd As Worksheet: Set wsProd = ThisWorkbook.Sheets("Production")
    Dim wsQR As Worksheet: Set wsQR = ThisWorkbook.Sheets(TEMPLATE_SHEET)

    Dim lastRow As Long: lastRow = wsProd.Cells(wsProd.Rows.Count, "BI").End(xlUp).Row
    Dim fdcList As Collection: Set fdcList = New Collection
    Dim fdcRowMap As Object: Set fdcRowMap = CreateObject("Scripting.Dictionary")

    Dim i As Long, fdc As String

    ' Build list of valid FDCs
    For i = START_ROW To lastRow
        fdc = Trim(wsProd.Cells(i, "BI").Value)
        If fdc <> "" And Left(fdc, 1) <> "-" Then
            If Not fdcRowMap.exists(fdc) Then
                fdcList.Add fdc
                fdcRowMap.Add fdc, i
            End If
        End If
    Next i

    If fdcList.Count = 0 Then
        MsgBox "No valid FDCs with QR URLs were found to export.", vbExclamation
        Exit Sub
    End If

    ' Clear all previous headers, QR formulas, and QR images
    Dim r As Long, c As Long, sh As Shape
    For r = 4 To 85
        For c = 3 To 19 Step 4
            On Error Resume Next
            wsQR.Cells(r, c).ClearContents
            wsQR.Cells(r + 1, c).ClearContents
            On Error GoTo 0
        Next c
    Next r

    For Each sh In wsQR.Shapes
        If sh.Type = msoPicture Then sh.Delete
    Next sh

    ' Fill QR labels
    For i = 1 To fdcList.Count
        fdc = fdcList(i)
        Dim prodRow As Long: prodRow = fdcRowMap(fdc)

        Dim slot As Long: slot = ((i - 1) Mod SLOT_COUNT)
        Dim pageGroup As Long: pageGroup = slot \ 20     ' 0, 1, 2
        Dim rowBlock As Long: rowBlock = (slot Mod 20) \ 4
        Dim colBlock As Long: colBlock = (slot Mod 4)

        Dim baseRow As Long
        Select Case pageGroup
            Case 0: baseRow = 4
            Case 1: baseRow = 32
            Case 2: baseRow = 60
        End Select

        Dim headerRow As Long: headerRow = baseRow + (rowBlock * 5)
        Dim qrRow As Long: qrRow = headerRow + 1
        Dim col As Long: col = 3 + (colBlock * 4)

        ' Write header and QR
        wsQR.Cells(headerRow, col).Formula = "=Production!BI" & prodRow
        wsQR.Cells(qrRow, col).Formula = "=Production!BT" & prodRow
    Next i

    ' Set correct print area
    Dim lastUsedRow As Long
    Select Case fdcList.Count
        Case Is <= 20: lastUsedRow = 29
        Case Is <= 40: lastUsedRow = 57
        Case Else: lastUsedRow = 85
    End Select

    With wsQR.PageSetup
        .printArea = "B3:R" & lastUsedRow
        .Zoom = 100
        .FitToPagesWide = False
        .FitToPagesTall = False
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .CenterVertically = False
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
    End With

    wsQR.ResetAllPageBreaks
    If fdcList.Count > 20 Then wsQR.HPageBreaks.Add Before:=wsQR.Rows(30)
    If fdcList.Count > 40 Then wsQR.HPageBreaks.Add Before:=wsQR.Rows(58)

    ' Export to PDF
    Dim exportPath As String
    exportPath = ThisWorkbook.Path & "\" & OUTPUT_NAME
    wsQR.Range("B3:R" & lastUsedRow).ExportAsFixedFormat Type:=xlTypePDF, fileName:=exportPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox fdcList.Count & " QR labels exported as PDF." & vbNewLine & exportPath, vbInformation
End Sub


