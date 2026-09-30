Attribute VB_Name = "modAETPoints"
Option Explicit

Public Sub AET_NumberPoints()
    On Error GoTo Failed
    Dim pointCount As Long
    Dim index As Long
    Dim markerRadius As Double
    Dim textHeight As Double
    Dim countValue As Long
    Dim location As Variant
    Dim marker As Object
    Dim label As Object

    countValue = AET_GetWholeNumber("Number of points to place:", 1)
    If countValue <= 0 Then Err.Raise vbObjectError + 1037, "AET", "Point count must be greater than zero."
    pointCount = countValue
    If pointCount > 500 Then Err.Raise vbObjectError + 1020, "AET", "Limit the operation to 500 points at a time."
    markerRadius = AET_GetPositiveNumber("Marker radius in drawing units:", 0.5)
    textHeight = AET_GetPositiveNumber("Label text height in drawing units:", 2.5)
    Call AET_EnsureLayer(AET_LAYER_POINTS, 3)
    ThisDrawing.StartUndoMark
    For index = 1 To pointCount
        location = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick setting-out point " & CStr(index) & ": ")
        Set marker = ThisDrawing.ModelSpace.AddCircle(location, markerRadius)
        marker.Layer = AET_LAYER_POINTS
        Set label = ThisDrawing.ModelSpace.AddText(CStr(index), location, textHeight)
        label.Layer = AET_LAYER_POINTS
        location = AET_Point3D(CDbl(location(0)) + markerRadius, CDbl(location(1)) + markerRadius, CDbl(location(2)))
        label.InsertionPoint = location
    Next index
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & CStr(pointCount) & " numbered point(s) created." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_NumberPoints"
End Sub

Public Sub AET_ExportPointsCsv()
    On Error GoTo Failed
    Dim selection As Object
    Dim entity As Object
    Dim position As Variant
    Dim filePath As String
    Dim fileNumber As Integer
    Dim pointCount As Long

    filePath = AET_GetText("CSV output path (for example C:\Temp\points.csv):", "points.csv")
    If Len(Dir$(filePath)) > 0 Then
        If MsgBox("File already exists. Overwrite " & filePath & "?", vbYesNo + vbQuestion, _
                  "Confirm CSV Export") <> vbYes Then Err.Raise vbObjectError + 1044, "AET", "Export cancelled."
    End If
    Set selection = AET_SelectEntities("Select point objects and block references to export: ")
    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, "id,x,y,z"
    For Each entity In selection
        Select Case entity.ObjectName
            Case "AcDbPoint"
                position = entity.Position
            Case "AcDbBlockReference"
                position = entity.InsertionPoint
            Case Else
                GoTo NextEntity
        End Select
        pointCount = pointCount + 1
        Print #fileNumber, """" & CStr(pointCount) & """," & _
            AET_CsvNumber(CDbl(position(0))) & "," & AET_CsvNumber(CDbl(position(1))) & "," & AET_CsvNumber(CDbl(position(2)))
NextEntity:
    Next entity
    Close #fileNumber
    ThisDrawing.Utility.Prompt vbCrLf & CStr(pointCount) & " point(s) exported to " & filePath & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    AET_ReportError "AET_ExportPointsCsv"
End Sub

Public Sub AET_ImportPointsCsv()
    On Error GoTo Failed
    Dim filePath As String
    Dim fileNumber As Integer
    Dim lineText As String
    Dim fields As Variant
    Dim point As Object
    Dim pointRows As Collection
    Dim row As Variant
    Dim imported As Long

    filePath = AET_GetText("CSV input path (format: id,x,y,z):", "points.csv")
    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    Set pointRows = New Collection
    If EOF(fileNumber) Then Err.Raise vbObjectError + 1027, "AET", "CSV file is empty; expected header id,x,y,z."
    Line Input #fileNumber, lineText
    If LCase$(Trim$(lineText)) <> "id,x,y,z" Then _
        Err.Raise vbObjectError + 1028, "AET", "First CSV row must be the header id,x,y,z."
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        If Len(Trim$(lineText)) > 0 Then
            fields = Split(lineText, ",")
            If UBound(fields) <> 3 Then Err.Raise vbObjectError + 1021, "AET", "CSV row must contain id,x,y,z: " & lineText
            If Len(Trim$(CStr(fields(0)))) = 0 Then Err.Raise vbObjectError + 1029, "AET", "CSV point ID is empty: " & lineText
            pointRows.Add Array(AET_ParseCsvCoordinate(CStr(fields(1)), lineText), _
                                AET_ParseCsvCoordinate(CStr(fields(2)), lineText), _
                                AET_ParseCsvCoordinate(CStr(fields(3)), lineText))
        End If
    Loop
    Close #fileNumber
    Call AET_EnsureLayer(AET_LAYER_POINTS, 3)
    ThisDrawing.StartUndoMark
    For Each row In pointRows
            Set point = ThisDrawing.ModelSpace.AddPoint(AET_Point3D(CDbl(row(0)), CDbl(row(1)), CDbl(row(2))))
            point.Layer = AET_LAYER_POINTS
            imported = imported + 1
    Next row
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & CStr(imported) & " point(s) imported from " & filePath & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_ImportPointsCsv"
End Sub

Private Function AET_ParseCsvCoordinate(ByVal coordinateText As String, ByVal rowText As String) As Double
    Dim valueText As String
    Dim index As Long
    Dim character As String
    Dim digitCount As Long
    Dim decimalCount As Long
    valueText = Trim$(coordinateText)
    If Len(valueText) = 0 Then GoTo InvalidCoordinate
    For index = 1 To Len(valueText)
        character = Mid$(valueText, index, 1)
        If character >= "0" And character <= "9" Then
            digitCount = digitCount + 1
        ElseIf character = "." Then
            decimalCount = decimalCount + 1
            If decimalCount > 1 Then GoTo InvalidCoordinate
        ElseIf (character = "-" Or character = "+") And index = 1 Then
        Else
            GoTo InvalidCoordinate
        End If
    Next index
    If digitCount = 0 Then GoTo InvalidCoordinate
    AET_ParseCsvCoordinate = Val(valueText)
    Exit Function
InvalidCoordinate:
    Err.Raise vbObjectError + 1022, "AET", "CSV coordinates must be signed decimal numbers using a period: " & rowText
End Function

Private Function AET_CsvNumber(ByVal coordinate As Double) As String
    AET_CsvNumber = Trim$(Str$(coordinate))
End Function

Public Sub AET_LabelLevels()
    On Error GoTo Failed
    Dim pointCount As Long
    Dim precision As Long
    Dim textHeight As Double
    Dim prefix As String
    Dim suffix As String
    Dim index As Long
    Dim location As Variant
    Dim label As Object
    Dim textValue As String
    Dim countValue As Long

    countValue = AET_GetWholeNumber("Number of spot levels to label:", 1)
    If countValue <= 0 Then Err.Raise vbObjectError + 1038, "AET", "Level count must be greater than zero."
    pointCount = countValue
    If pointCount > 500 Then Err.Raise vbObjectError + 1023, "AET", "Limit the operation to 500 levels at a time."
    precision = AET_GetWholeNumber("Decimal places (0-8):", 2)
    If precision < 0 Or precision > 8 Then Err.Raise vbObjectError + 1024, "AET", "Precision must be between 0 and 8."
    prefix = InputBox("Optional level prefix:", "AutoCAD Engineering Toolkit", "EL ")
    suffix = InputBox("Optional level suffix:", "AutoCAD Engineering Toolkit", "")
    textHeight = AET_GetPositiveNumber("Text height in drawing units:", 2.5)
    Call AET_EnsureLayer(AET_LAYER_LEVELS, 2)
    ThisDrawing.StartUndoMark
    For index = 1 To pointCount
        location = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick spot level point " & CStr(index) & ": ")
        If precision = 0 Then
            textValue = prefix & Format$(CDbl(location(2)), "0") & suffix
        Else
            textValue = prefix & Format$(CDbl(location(2)), "0." & String$(precision, "0")) & suffix
        End If
        Set label = ThisDrawing.ModelSpace.AddText(textValue, location, textHeight)
        label.Layer = AET_LAYER_LEVELS
    Next index
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & CStr(pointCount) & " spot level(s) labeled. Elevations use the picked point Z values." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_LabelLevels"
End Sub

Public Sub AET_LabelCoordinates()
    On Error GoTo Failed
    Dim pointCount As Long
    Dim index As Long
    Dim textHeight As Double
    Dim location As Variant
    Dim label As Object
    Dim labelText As String
    pointCount = AET_GetWholeNumber("Number of coordinate labels:", 1)
    If pointCount <= 0 Or pointCount > 500 Then _
        Err.Raise vbObjectError + 1039, "AET", "Choose between 1 and 500 coordinate labels."
    textHeight = AET_GetPositiveNumber("Text height in drawing units:", 2.5)
    Call AET_EnsureLayer(AET_LAYER_POINTS, 3)
    ThisDrawing.StartUndoMark
    For index = 1 To pointCount
        location = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick coordinate label point " & CStr(index) & ": ")
        labelText = "X " & AET_CsvNumber(CDbl(location(0))) & "  Y " & _
                    AET_CsvNumber(CDbl(location(1))) & "  Z " & AET_CsvNumber(CDbl(location(2)))
        Set label = ThisDrawing.ModelSpace.AddText(labelText, location, textHeight)
        label.Layer = AET_LAYER_POINTS
    Next index
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & CStr(pointCount) & " coordinate label(s) created." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_LabelCoordinates"
End Sub

Public Sub AET_ChainageOffset()
    On Error GoTo Failed
    Dim alignment As Object
    Dim pickedPoint As Variant
    Dim queryPoint As Variant
    Dim closest As Variant
    Dim beforePoint As Variant
    Dim afterPoint As Variant
    Dim startPoint As Variant
    Dim endPoint As Variant
    Dim distanceAlong As Double
    Dim totalLength As Double
    Dim sampleDistance As Double
    Dim tangentX As Double
    Dim tangentY As Double
    Dim offsetX As Double
    Dim offsetY As Double
    Dim offsetValue As Double
    Dim crossValue As Double

    ThisDrawing.Utility.GetEntity alignment, pickedPoint, vbCrLf & "Select a 2D line or polyline alignment: "
    If alignment.ObjectName <> "AcDbPolyline" And alignment.ObjectName <> "AcDb2dPolyline" And _
       alignment.ObjectName <> "AcDbLine" Then
        Err.Raise vbObjectError + 1025, "AET", "Select a line or 2D polyline."
    End If
    queryPoint = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick a point to measure from the alignment: ")
    closest = alignment.GetClosestPointTo(queryPoint, False)
    totalLength = AET_EntityLength(alignment)
    If alignment.ObjectName = "AcDbLine" Then
        startPoint = alignment.StartPoint
        endPoint = alignment.EndPoint
        distanceAlong = Sqr((CDbl(closest(0)) - CDbl(startPoint(0))) ^ 2 + _
                            (CDbl(closest(1)) - CDbl(startPoint(1))) ^ 2 + _
                            (CDbl(closest(2)) - CDbl(startPoint(2))) ^ 2)
        tangentX = CDbl(endPoint(0)) - CDbl(startPoint(0))
        tangentY = CDbl(endPoint(1)) - CDbl(startPoint(1))
    Else
        distanceAlong = alignment.GetDistAtPoint(closest)
        sampleDistance = totalLength / 10000#
        If sampleDistance < 0.000001 Then sampleDistance = 0.000001
        If distanceAlong > sampleDistance Then
            beforePoint = alignment.GetPointAtDist(distanceAlong - sampleDistance)
        Else
            beforePoint = alignment.GetPointAtDist(0#)
        End If
        If distanceAlong + sampleDistance < totalLength Then
            afterPoint = alignment.GetPointAtDist(distanceAlong + sampleDistance)
        Else
            afterPoint = alignment.GetPointAtDist(totalLength)
        End If
        tangentX = CDbl(afterPoint(0)) - CDbl(beforePoint(0))
        tangentY = CDbl(afterPoint(1)) - CDbl(beforePoint(1))
    End If
    offsetX = CDbl(queryPoint(0)) - CDbl(closest(0))
    offsetY = CDbl(queryPoint(1)) - CDbl(closest(1))
    offsetValue = Sqr(offsetX * offsetX + offsetY * offsetY)
    crossValue = tangentX * offsetY - tangentY * offsetX
    If crossValue < 0# Then offsetValue = -offsetValue
    ThisDrawing.Utility.Prompt vbCrLf & "Chainage=" & Format$(distanceAlong, "0.000") & _
        " drawing units; signed offset=" & Format$(offsetValue, "0.000") & _
        " drawing units (left is positive)." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_ChainageOffset"
End Sub

Public Sub AET_DrawGrid()
    On Error GoTo Failed
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    Dim spacing As Double
    Dim xValue As Double
    Dim yValue As Double
    Dim lineObject As Object
    Dim index As Long
    Dim xMin As Double
    Dim xMax As Double
    Dim yMin As Double
    Dim yMax As Double
    Dim plannedLines As Long

    firstPoint = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick first grid corner: ")
    secondPoint = ThisDrawing.Utility.GetPoint(firstPoint, vbCrLf & "Pick opposite grid corner: ")
    spacing = AET_GetPositiveNumber("Grid spacing in drawing units:", 10#)
    xMin = CDbl(firstPoint(0)): xMax = CDbl(secondPoint(0))
    yMin = CDbl(firstPoint(1)): yMax = CDbl(secondPoint(1))
    If xMax < xMin Then SwapValues xMin, xMax
    If yMax < yMin Then SwapValues yMin, yMax
    If xMax = xMin Or yMax = yMin Then Err.Raise vbObjectError + 1042, "AET", "Grid corners must define a nonzero area."
    plannedLines = Int((xMax - xMin) / spacing + 0.000001) + 1 + _
                   Int((yMax - yMin) / spacing + 0.000001) + 1
    If plannedLines > 2000 Then
        Err.Raise vbObjectError + 1026, "AET", "Grid would exceed 2,000 lines. Increase spacing or reduce the extents."
    End If
    Call AET_EnsureLayer(AET_LAYER_GRID, 8, "Dashed")
    ThisDrawing.StartUndoMark
    xValue = xMin
    Do While xValue <= xMax + spacing * 0.000001
        Set lineObject = ThisDrawing.ModelSpace.AddLine(AET_Point3D(xValue, yMin, CDbl(firstPoint(2))), _
                                                        AET_Point3D(xValue, yMax, CDbl(firstPoint(2))))
        lineObject.Layer = AET_LAYER_GRID
        xValue = xValue + spacing
        index = index + 1
    Loop
    yValue = yMin
    Do While yValue <= yMax + spacing * 0.000001
        Set lineObject = ThisDrawing.ModelSpace.AddLine(AET_Point3D(xMin, yValue, CDbl(firstPoint(2))), _
                                                        AET_Point3D(xMax, yValue, CDbl(firstPoint(2))))
        lineObject.Layer = AET_LAYER_GRID
        yValue = yValue + spacing
        index = index + 1
    Loop
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & CStr(index) & " grid line(s) drawn." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_DrawGrid"
End Sub

Private Sub SwapValues(ByRef firstValue As Double, ByRef secondValue As Double)
    Dim temporary As Double
    temporary = firstValue
    firstValue = secondValue
    secondValue = temporary
End Sub
