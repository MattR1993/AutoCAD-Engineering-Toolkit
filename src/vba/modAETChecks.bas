Attribute VB_Name = "modAETChecks"
Option Explicit

Public Sub AET_RunChecks()
    On Error GoTo Failed
    Dim coordinates(0 To 7) As Double
    Dim testPolyline As Object
    Dim measuredArea As Double
    Dim passed As Boolean

    coordinates(0) = 0#: coordinates(1) = 0#
    coordinates(2) = 4#: coordinates(3) = 0#
    coordinates(4) = 4#: coordinates(5) = 3#
    coordinates(6) = 0#: coordinates(7) = 3#
    Set testPolyline = ThisDrawing.ModelSpace.AddLightWeightPolyline(coordinates)
    testPolyline.Closed = True
    passed = AET_TryGetArea(testPolyline, measuredArea)
    If Not passed Or Abs(measuredArea - 12#) > 0.000001 Then _
        Err.Raise vbObjectError + 1050, "AET", "Rectangle area check failed; expected 12."
    If Abs(AET_EntityLength(testPolyline) - 14#) > 0.000001 Then _
        Err.Raise vbObjectError + 1051, "AET", "Rectangle perimeter check failed; expected 14."
    testPolyline.Delete
    Set testPolyline = Nothing
    ThisDrawing.Utility.Prompt vbCrLf & "AET checks passed: closed rectangle area = 12, perimeter = 14." & vbCrLf
    Exit Sub
Failed:
    On Error Resume Next
    If Not testPolyline Is Nothing Then testPolyline.Delete
    On Error GoTo 0
    AET_ReportError "AET_RunChecks"
End Sub
