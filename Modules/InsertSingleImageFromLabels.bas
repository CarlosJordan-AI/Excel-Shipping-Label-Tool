Attribute VB_Name = "Module3"
Sub InsertSingleImageFromLabels_V1()
    Dim wsLabels As Worksheet, wsShipping As Worksheet
    Dim searchValue As String
    Dim foundCell As Range
    Dim shp As Shape, clonedShape As Shape
    Dim rngTarget As Range
    Dim targetRow As Long

    Set wsLabels = ThisWorkbook.Sheets("Labels + Carriers")
    Set wsShipping = ThisWorkbook.Sheets("shipping label template")

    ' Get carrier name from D22 (updated)
    searchValue = Trim(wsShipping.Range("D22").Value)
    If searchValue = "" Then Exit Sub

    ' Clear any previous images in target range D23:G26
    For Each shp In wsShipping.Shapes
        If Not Intersect(shp.TopLeftCell, wsShipping.Range("D23:G26")) Is Nothing Then
            shp.Delete
        End If
    Next shp

    ' Find the carrier row in "Labels + Carriers"
    Set foundCell = wsLabels.Range("C:C").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then Exit Sub
    targetRow = foundCell.Row

    ' Find and duplicate the shape in that row (usually in column D)
    For Each shp In wsLabels.Shapes
        If shp.TopLeftCell.Row = targetRow And shp.TopLeftCell.Column = 4 Then
            Set clonedShape = shp.Duplicate
            Exit For
        End If
    Next shp

    If clonedShape Is Nothing Then Exit Sub

    ' Move duplicated shape to shipping label template
    With clonedShape
        .Cut
        wsShipping.Paste
        Set clonedShape = wsShipping.Shapes(wsShipping.Shapes.Count)
    End With

    ' Resize and center the image within D23:G26
    Set rngTarget = wsShipping.Range("D23:G26")
    With clonedShape
        .LockAspectRatio = msoTrue
        If .Width > rngTarget.Width Then .Width = rngTarget.Width
        If .Height > rngTarget.Height Then .Height = rngTarget.Height
        .Top = rngTarget.Top + (rngTarget.Height - .Height) / 2
        .Left = rngTarget.Left + (rngTarget.Width - .Width) / 2
    End With
End Sub


