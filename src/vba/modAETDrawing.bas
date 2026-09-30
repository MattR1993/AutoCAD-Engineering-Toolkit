Attribute VB_Name = "modAETDrawing"
Option Explicit

Public Sub AET_DrawConstructionLine()
    On Error GoTo Failed
    Dim firstPoint As Variant
    Dim secondPoint As Variant
    Dim direction(0 To 2) As Double
    Dim constructionLine As Object
    firstPoint = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick construction-line origin: ")
    secondPoint = ThisDrawing.Utility.GetPoint(firstPoint, vbCrLf & "Pick a point defining its direction: ")
    direction(0) = CDbl(secondPoint(0)) - CDbl(firstPoint(0))
    direction(1) = CDbl(secondPoint(1)) - CDbl(firstPoint(1))
    direction(2) = CDbl(secondPoint(2)) - CDbl(firstPoint(2))
    If direction(0) = 0# And direction(1) = 0# And direction(2) = 0# Then _
        Err.Raise vbObjectError + 1040, "AET", "Direction points must be different."
    Call AET_EnsureLayer(AET_LAYER_HELPERS, 4, "Center")
    Set constructionLine = ThisDrawing.ModelSpace.AddXline(firstPoint, direction)
    constructionLine.Layer = AET_LAYER_HELPERS
    ThisDrawing.Utility.Prompt vbCrLf & "Construction line created on " & AET_LAYER_HELPERS & "." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_DrawConstructionLine"
End Sub

Public Sub AET_OffsetSelectedGeometry()
    On Error GoTo Failed
    Dim entity As Object
    Dim pickedPoint As Variant
    Dim offsetDistance As Double
    Dim offsetObjects As Variant
    Dim newEntity As Object
    Dim index As Long
    ThisDrawing.Utility.GetEntity entity, pickedPoint, vbCrLf & "Select a line or polyline to offset: "
    If entity.ObjectName <> "AcDbLine" And entity.ObjectName <> "AcDbPolyline" And _
       entity.ObjectName <> "AcDb2dPolyline" Then
        Err.Raise vbObjectError + 1041, "AET", "Only lines and 2D polylines can be offset."
    End If
    offsetDistance = AET_GetPositiveNumber("Offset distance in drawing units:", 1#)
    Call AET_EnsureLayer(AET_LAYER_HELPERS, 4, "Center")
    ThisDrawing.StartUndoMark
    offsetObjects = entity.Offset(offsetDistance)
    For index = LBound(offsetObjects) To UBound(offsetObjects)
        Set newEntity = offsetObjects(index)
        newEntity.Layer = AET_LAYER_HELPERS
    Next index
    ThisDrawing.EndUndoMark
    ThisDrawing.Utility.Prompt vbCrLf & "Offset object(s) created on " & AET_LAYER_HELPERS & "." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    ThisDrawing.EndUndoMark
    On Error GoTo 0
    AET_ReportError "AET_OffsetSelectedGeometry"
End Sub

Public Sub AET_DrawFoundationOutline()
    On Error GoTo Failed
    Dim origin As Variant
    Dim lengthValue As Double
    Dim widthValue As Double
    Dim coordinates(0 To 7) As Double
    Dim outline As Object
    origin = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick foundation lower-left corner: ")
    lengthValue = AET_GetPositiveNumber("Foundation length in drawing units (X):", 5#)
    widthValue = AET_GetPositiveNumber("Foundation width in drawing units (Y):", 3#)
    coordinates(0) = CDbl(origin(0)): coordinates(1) = CDbl(origin(1))
    coordinates(2) = CDbl(origin(0)) + lengthValue: coordinates(3) = CDbl(origin(1))
    coordinates(4) = CDbl(origin(0)) + lengthValue: coordinates(5) = CDbl(origin(1)) + widthValue
    coordinates(6) = CDbl(origin(0)): coordinates(7) = CDbl(origin(1)) + widthValue
    Call AET_EnsureLayer(AET_LAYER_HELPERS, 4, "Center")
    Set outline = ThisDrawing.ModelSpace.AddLightWeightPolyline(coordinates)
    outline.Closed = True
    outline.Elevation = CDbl(origin(2))
    outline.Layer = AET_LAYER_HELPERS
    ThisDrawing.Utility.Prompt vbCrLf & "Rectangular foundation outline created; dimensions are drawing units." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_DrawFoundationOutline"
End Sub

Public Sub AET_DrawDrainageNote()
    On Error GoTo Failed
    Dim insertionPoint As Variant
    Dim noteText As String
    Dim textHeight As Double
    Dim note As Object
    noteText = AET_GetText("Drainage note (for example: 150 mm dia. drain, fall 1:100):", "Drain to invert level")
    textHeight = AET_GetPositiveNumber("Text height in drawing units:", 2.5)
    insertionPoint = ThisDrawing.Utility.GetPoint(, vbCrLf & "Pick note insertion point: ")
    Call AET_EnsureLayer(AET_LAYER_NOTES, 7)
    Set note = ThisDrawing.ModelSpace.AddText(noteText, insertionPoint, textHeight)
    note.Layer = AET_LAYER_NOTES
    ThisDrawing.Utility.Prompt vbCrLf & "Drainage annotation added to " & AET_LAYER_NOTES & "." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_DrawDrainageNote"
End Sub

Public Sub AET_AuditDrawing()
    On Error GoTo Failed
    If MsgBox("Run AutoCAD AUDIT with repair enabled? This may modify the drawing database.", _
              vbYesNo + vbExclamation, "Confirm Drawing Audit") <> vbYes Then Exit Sub
    ThisDrawing.SendCommand "_.AUDIT" & vbCr & "_Y" & vbCr
    ThisDrawing.Utility.Prompt vbCrLf & "AUDIT started. Review AutoCAD's command-line results." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_AuditDrawing"
End Sub

Public Sub AET_CleanupDrawing()
    On Error GoTo Failed
    If MsgBox("Start AutoCAD PURGE? Purging removes unused definitions and is not generally undoable.", _
              vbYesNo + vbExclamation, "Confirm Drawing Cleanup") <> vbYes Then Exit Sub
    ThisDrawing.SendCommand "_.-PURGE" & vbCr
    ThisDrawing.Utility.Prompt vbCrLf & _
        "Complete PURGE at the command line. Review each prompt; press Esc to cancel." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_CleanupDrawing"
End Sub
