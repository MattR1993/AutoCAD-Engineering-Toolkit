Attribute VB_Name = "modAETCore"
Option Explicit

Public Const AET_LAYER_POINTS As String = "AET-POINTS"
Public Const AET_LAYER_LEVELS As String = "AET-LEVELS"
Public Const AET_LAYER_GRID As String = "AET-GRID"
Public Const AET_LAYER_NOTES As String = "AET-NOTES"
Public Const AET_LAYER_HELPERS As String = "AET-HELPERS"

Public Sub AET_Menu()
    MsgBox "AutoCAD Engineering Toolkit" & vbCrLf & vbCrLf & _
           "Layers: AET_SetupLayers, AET_SetCurrentLayer" & vbCrLf & _
           "Geometry: AET_CalculateArea, AET_GeometryReport" & vbCrLf & _
           "Setting out: AET_NumberPoints, AET_LabelCoordinates, AET_ChainageOffset, AET_DrawGrid" & vbCrLf & _
           "Coordinates: AET_ExportPointsCsv, AET_ImportPointsCsv" & vbCrLf & _
           "Levels and volumes: AET_LabelLevels, AET_CalculateVolume" & vbCrLf & _
           "Drawing helpers: AET_DrawConstructionLine, AET_DrawDrainageNote, AET_CleanupDrawing, AET_AuditDrawing", _
           vbInformation, "AET Commands"
End Sub

Public Function AET_GetPositiveNumber(ByVal prompt As String, ByVal defaultValue As Double) As Double
    Dim answer As String
    answer = InputBox(prompt, "AutoCAD Engineering Toolkit", CStr(defaultValue))
    If Len(answer) = 0 Then Err.Raise vbObjectError + 1000, "AET", "Operation cancelled."
    If Not IsNumeric(answer) Then Err.Raise vbObjectError + 1001, "AET", "Enter a numeric value."
    AET_GetPositiveNumber = CDbl(answer)
    If AET_GetPositiveNumber <= 0# Then Err.Raise vbObjectError + 1002, "AET", "Enter a value greater than zero."
End Function

Public Function AET_GetNumber(ByVal prompt As String, ByVal defaultValue As Double) As Double
    Dim answer As String
    answer = InputBox(prompt, "AutoCAD Engineering Toolkit", CStr(defaultValue))
    If Len(answer) = 0 Then Err.Raise vbObjectError + 1003, "AET", "Operation cancelled."
    If Not IsNumeric(answer) Then Err.Raise vbObjectError + 1004, "AET", "Enter a numeric value."
    AET_GetNumber = CDbl(answer)
End Function

Public Function AET_GetWholeNumber(ByVal prompt As String, ByVal defaultValue As Long) As Long
    Dim value As Double
    value = AET_GetNumber(prompt, CDbl(defaultValue))
    If value <> Fix(value) Then Err.Raise vbObjectError + 1006, "AET", "Enter a whole number."
    AET_GetWholeNumber = CLng(value)
End Function

Public Function AET_GetText(ByVal prompt As String, ByVal defaultValue As String) As String
    AET_GetText = InputBox(prompt, "AutoCAD Engineering Toolkit", defaultValue)
    If Len(AET_GetText) = 0 Then Err.Raise vbObjectError + 1005, "AET", "Operation cancelled or empty value."
End Function

Public Function AET_Point3D(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    AET_Point3D = Array(x, y, z)
End Function

Public Sub AET_ReportError(ByVal procedureName As String)
    If Err.Number <> 0 Then
        If Err.Number = vbObjectError + 1000 Or Err.Number = vbObjectError + 1003 Or _
           Err.Number = vbObjectError + 1005 Then
            ThisDrawing.Utility.Prompt vbCrLf & Err.Description & vbCrLf
        Else
            ThisDrawing.Utility.Prompt vbCrLf & procedureName & " failed: " & Err.Description & vbCrLf
        End If
    End If
End Sub

Public Function AET_EnsureLayer(ByVal layerName As String, ByVal colorIndex As Integer, _
                                Optional ByVal lineTypeName As String = "Continuous") As Object
    Dim layer As Object
    On Error Resume Next
    Set layer = ThisDrawing.Layers.Item(layerName)
    On Error GoTo 0
    If layer Is Nothing Then Set layer = ThisDrawing.Layers.Add(layerName)

    layer.Color = colorIndex
    On Error Resume Next
    layer.Linetype = lineTypeName
    If Err.Number <> 0 Then
        Err.Clear
        ThisDrawing.Linetypes.Load lineTypeName, "acad.lin"
        layer.Linetype = lineTypeName
    End If
    On Error GoTo 0
    Set AET_EnsureLayer = layer
End Function
