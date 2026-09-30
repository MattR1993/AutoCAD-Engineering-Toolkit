Attribute VB_Name = "modAETGeometry"
Option Explicit

Public Sub AET_GeometryReport()
    On Error GoTo Failed
    Dim selection As Object
    Dim entity As Object
    Dim lengthValue As Double
    Dim areaValue As Double
    Dim minPoint As Variant
    Dim maxPoint As Variant
    Dim minExt As Variant
    Dim maxExt As Variant
    Dim count As Long

    Set selection = AET_SelectEntities("Select lines and polylines for a geometry report: ")
    If selection.Count = 0 Then GoTo Finished
    For Each entity In selection
        count = count + 1
        areaValue = 0#
        If entity.ObjectName <> "AcDbRegion" Then lengthValue = AET_EntityLength(entity)
        If AET_TryGetArea(entity, areaValue) Then
            If entity.ObjectName = "AcDbRegion" Then
                ThisDrawing.Utility.Prompt vbCrLf & entity.ObjectName & " area=" & CStr(areaValue)
            Else
                ThisDrawing.Utility.Prompt vbCrLf & entity.ObjectName & " length/perimeter=" & _
                    CStr(lengthValue) & ", area=" & CStr(areaValue)
            End If
        Else
            ThisDrawing.Utility.Prompt vbCrLf & entity.ObjectName & " length=" & CStr(lengthValue)
        End If
        entity.GetBoundingBox minExt, maxExt
        minPoint = minExt
        maxPoint = maxExt
        ThisDrawing.Utility.Prompt ", bounds=(" & CStr(minPoint(0)) & ", " & CStr(minPoint(1)) & _
            ", " & CStr(maxPoint(0)) & ", " & CStr(maxPoint(1)) & ")"
    Next entity
    ThisDrawing.Utility.Prompt vbCrLf & CStr(count) & " object(s) reported in drawing units." & vbCrLf
Finished:
    Exit Sub
Failed:
    AET_ReportError "AET_GeometryReport"
End Sub

Public Sub AET_CalculateArea()
    On Error GoTo Failed
    Dim entity As Object
    Dim pickedPoint As Variant
    Dim areaValue As Double
    ThisDrawing.Utility.GetEntity entity, pickedPoint, vbCrLf & "Select a closed polyline, circle, or region: "
    If Not AET_TryGetArea(entity, areaValue) Then
        Err.Raise vbObjectError + 1010, "AET", "The selected object has no supported area. Use a closed polyline, circle, or region."
    End If
    ThisDrawing.Utility.Prompt vbCrLf & "Area=" & CStr(areaValue) & " square drawing units."
    If entity.ObjectName = "AcDbRegion" Then
        ThisDrawing.Utility.Prompt vbCrLf
    Else
        ThisDrawing.Utility.Prompt " Perimeter/length=" & CStr(AET_EntityLength(entity)) & " drawing units." & vbCrLf
    End If
    Exit Sub
Failed:
    AET_ReportError "AET_CalculateArea"
End Sub

Public Function AET_SelectEntities(ByVal prompt As String) As Object
    Dim selection As Object
    On Error Resume Next
    ThisDrawing.SelectionSets.Item("AET_TEMP").Delete
    On Error GoTo 0
    Set selection = ThisDrawing.SelectionSets.Add("AET_TEMP")
    ThisDrawing.Utility.Prompt vbCrLf & prompt
    selection.SelectOnScreen
    Set AET_SelectEntities = selection
End Function

Public Function AET_EntityLength(ByVal entity As Object) As Double
    Dim startPoint As Variant
    Dim endPoint As Variant
    Select Case entity.ObjectName
        Case "AcDbLine"
            startPoint = entity.StartPoint
            endPoint = entity.EndPoint
            AET_EntityLength = Sqr((endPoint(0) - startPoint(0)) ^ 2 + _
                                   (endPoint(1) - startPoint(1)) ^ 2 + _
                                   (endPoint(2) - startPoint(2)) ^ 2)
        Case "AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"
            AET_EntityLength = entity.Length
        Case "AcDbArc"
            AET_EntityLength = entity.ArcLength
        Case "AcDbCircle"
            AET_EntityLength = entity.Circumference
        Case Else
            Err.Raise vbObjectError + 1011, "AET", "Unsupported object for length: " & entity.ObjectName
    End Select
End Function

Public Function AET_TryGetArea(ByVal entity As Object, ByRef areaValue As Double) As Boolean
    On Error GoTo Unsupported
    Select Case entity.ObjectName
        Case "AcDbPolyline", "AcDb2dPolyline", "AcDbCircle", "AcDbRegion"
            If entity.ObjectName = "AcDbPolyline" Or entity.ObjectName = "AcDb2dPolyline" Then
                If Not entity.Closed Then Exit Function
            End If
            areaValue = entity.Area
            AET_TryGetArea = True
    End Select
    Exit Function
Unsupported:
    AET_TryGetArea = False
End Function
