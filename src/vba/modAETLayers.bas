Attribute VB_Name = "modAETLayers"
Option Explicit

Public Sub AET_SetupLayers()
    On Error GoTo Failed
    Call AET_EnsureLayer(AET_LAYER_POINTS, 3)
    Call AET_EnsureLayer(AET_LAYER_LEVELS, 2)
    Call AET_EnsureLayer(AET_LAYER_GRID, 8, "Dashed")
    Call AET_EnsureLayer(AET_LAYER_NOTES, 7)
    Call AET_EnsureLayer(AET_LAYER_HELPERS, 4, "Center")
    ThisDrawing.Utility.Prompt vbCrLf & "AET layers are ready." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_SetupLayers"
End Sub

Public Sub AET_SetCurrentLayer()
    On Error GoTo Failed
    Dim layerName As String
    Dim layer As Object
    layerName = AET_GetText("Layer name to make current:", AET_LAYER_NOTES)
    On Error Resume Next
    Set layer = ThisDrawing.Layers.Item(layerName)
    On Error GoTo Failed
    If layer Is Nothing Then
        Set layer = ThisDrawing.Layers.Add(layerName)
        layer.Color = 7
    End If
    ThisDrawing.ActiveLayer = layer
    ThisDrawing.Utility.Prompt vbCrLf & "Current layer: " & layerName & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_SetCurrentLayer"
End Sub

Public Sub AET_IsolateSelectedLayer()
    On Error GoTo Failed
    ThisDrawing.Utility.Prompt vbCrLf & "Select an object when AutoCAD prompts. Run LAYUNISO to restore the prior layer state." & vbCrLf
    ThisDrawing.SendCommand "_.LAYISO" & vbCr
    Exit Sub
Failed:
    AET_ReportError "AET_IsolateSelectedLayer"
End Sub

Public Sub AET_ThawAllLayers()
    On Error GoTo Failed
    Dim layer As Object
    If MsgBox("Thaw every layer except reserved layer 0 and Defpoints?", _
              vbYesNo + vbQuestion, "Confirm Thaw Layers") <> vbYes Then Exit Sub
    For Each layer In ThisDrawing.Layers
        If Not layer.Name = "0" And Not layer.Name = "Defpoints" Then layer.Freeze = False
    Next layer
    ThisDrawing.Utility.Prompt vbCrLf & "All non-reserved layers thawed." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_ThawAllLayers"
End Sub
