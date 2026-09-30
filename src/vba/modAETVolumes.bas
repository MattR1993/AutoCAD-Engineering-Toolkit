Attribute VB_Name = "modAETVolumes"
Option Explicit

Private Const AET_GRID_SUBDIVISIONS As Long = 4

Public Sub AET_CalculateVolume()
    On Error GoTo Failed
    Dim boundary As Object
    Dim pickedPoint As Variant
    Dim coordinates As Variant
    Dim boundaryNormal As Variant
    Dim vertexCount As Long
    Dim vertices() As Double
    Dim minExt As Variant
    Dim maxExt As Variant
    Dim minPoint As Variant
    Dim maxPoint As Variant
    Dim spacing As Double
    Dim baseLevel As Double
    Dim comparisonLevel As Double
    Dim thickness As Double
    Dim cellWidth As Double
    Dim sampleArea As Double
    Dim coveredSamples As Long
    Dim totalSamples As Long
    Dim volume As Double
    Dim xIndex As Long
    Dim yIndex As Long
    Dim sampleX As Double
    Dim sampleY As Double
    Dim label As String

    ThisDrawing.Utility.GetEntity boundary, pickedPoint, vbCrLf & "Select a closed, straight-segment 2D polyline boundary: "
    If boundary.ObjectName <> "AcDbPolyline" Then
        Err.Raise vbObjectError + 1030, "AET", "Volume boundaries must be closed lightweight 2D polylines."
    End If
    If Not boundary.Closed Then
        Err.Raise vbObjectError + 1030, "AET", "Volume boundaries must be closed lightweight 2D polylines."
    End If
    boundaryNormal = boundary.Normal
    If Abs(CDbl(boundaryNormal(0))) > 0.000001 Or Abs(CDbl(boundaryNormal(1))) > 0.000001 Or _
       CDbl(boundaryNormal(2)) < 0.999999 Then
        Err.Raise vbObjectError + 1043, "AET", "Volume boundary must lie in world XY with a positive Z normal."
    End If
    coordinates = boundary.Coordinates
    vertexCount = (UBound(coordinates) + 1) \ 2
    If vertexCount < 3 Or vertexCount > 1000 Then
        Err.Raise vbObjectError + 1031, "AET", "Boundary must have 3 to 1,000 vertices."
    End If
    ReDim vertices(0 To vertexCount - 1, 0 To 1)
    For xIndex = 0 To vertexCount - 1
        If Abs(CDbl(boundary.GetBulge(xIndex))) > 0.0000001 Then
            Err.Raise vbObjectError + 1032, "AET", "Curved boundary segments are not supported by the grid approximation."
        End If
        vertices(xIndex, 0) = CDbl(coordinates(xIndex * 2))
        vertices(xIndex, 1) = CDbl(coordinates(xIndex * 2 + 1))
    Next xIndex
    If AET_PolygonSelfIntersects(vertices, vertexCount) Then
        Err.Raise vbObjectError + 1033, "AET", "Boundary appears self-intersecting; use a simple closed outline."
    End If
    boundary.GetBoundingBox minExt, maxExt
    minPoint = minExt: maxPoint = maxExt
    If MsgBox("Calculate a positive stockpile/fill if the second level is above the base, or cut if below?", _
              vbYesNo + vbQuestion, "AET Grid Volume") <> vbYes Then Err.Raise vbObjectError + 1034, "AET", "Operation cancelled."
    baseLevel = AET_GetNumber("Base/existing average level:", 0#)
    comparisonLevel = AET_GetNumber("Stockpile top/design average level:", 1#)
    spacing = AET_GetPositiveNumber("Grid cell spacing in drawing units:", 5#)
    thickness = Abs(comparisonLevel - baseLevel)
    If thickness = 0# Then
        ThisDrawing.Utility.Prompt vbCrLf & "Volume is zero; the two input levels are equal." & vbCrLf
        Exit Sub
    End If

    cellWidth = spacing / AET_GRID_SUBDIVISIONS
    sampleArea = cellWidth * cellWidth
    xIndex = 0
    Do While CDbl(minPoint(0)) + xIndex * cellWidth < CDbl(maxPoint(0))
        yIndex = 0
        Do While CDbl(minPoint(1)) + yIndex * cellWidth < CDbl(maxPoint(1))
            sampleX = CDbl(minPoint(0)) + (xIndex + 0.5) * cellWidth
            sampleY = CDbl(minPoint(1)) + (yIndex + 0.5) * cellWidth
            If AET_PointInPolygon(sampleX, sampleY, vertices, vertexCount) Then coveredSamples = coveredSamples + 1
            totalSamples = totalSamples + 1
            yIndex = yIndex + 1
            If totalSamples > 500000 Then Err.Raise vbObjectError + 1035, "AET", "Grid exceeds 500,000 samples; increase spacing."
        Loop
        xIndex = xIndex + 1
    Loop
    If coveredSamples = 0 Then Err.Raise vbObjectError + 1036, "AET", "Grid resolution found no area inside the boundary; reduce spacing."
    volume = coveredSamples * sampleArea * thickness
    If comparisonLevel > baseLevel Then
        label = "Stockpile/fill"
    Else
        label = "Cut"
    End If
    ThisDrawing.Utility.Prompt vbCrLf & label & " volume (grid approximation)=" & Format$(volume, "0.000") & _
        " cubic drawing units; sampled plan area=" & Format$(coveredSamples * sampleArea, "0.000") & _
        " square drawing units; samples=" & CStr(totalSamples) & "." & vbCrLf
    ThisDrawing.Utility.Prompt "Assumes constant average thickness between the two supplied levels; verify against survey/design surfaces." & vbCrLf
    Exit Sub
Failed:
    AET_ReportError "AET_CalculateVolume"
End Sub

Private Function AET_PointInPolygon(ByVal xValue As Double, ByVal yValue As Double, _
                                    ByRef vertices() As Double, ByVal vertexCount As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim intersects As Boolean
    j = vertexCount - 1
    For i = 0 To vertexCount - 1
        If ((vertices(i, 1) > yValue) <> (vertices(j, 1) > yValue)) Then
            If xValue < (vertices(j, 0) - vertices(i, 0)) * (yValue - vertices(i, 1)) / _
                        (vertices(j, 1) - vertices(i, 1)) + vertices(i, 0) Then
                intersects = Not intersects
            End If
        End If
        j = i
    Next i
    AET_PointInPolygon = intersects
End Function

Private Function AET_PolygonSelfIntersects(ByRef vertices() As Double, ByVal vertexCount As Long) As Boolean
    Dim firstEdge As Long
    Dim secondEdge As Long
    Dim nextFirst As Long
    Dim nextSecond As Long
    For firstEdge = 0 To vertexCount - 1
        nextFirst = (firstEdge + 1) Mod vertexCount
        For secondEdge = firstEdge + 1 To vertexCount - 1
            nextSecond = (secondEdge + 1) Mod vertexCount
            If secondEdge <> firstEdge And secondEdge <> nextFirst And nextSecond <> firstEdge Then
                If AET_SegmentsCross(vertices(firstEdge, 0), vertices(firstEdge, 1), _
                    vertices(nextFirst, 0), vertices(nextFirst, 1), vertices(secondEdge, 0), _
                    vertices(secondEdge, 1), vertices(nextSecond, 0), vertices(nextSecond, 1)) Then
                    AET_PolygonSelfIntersects = True
                    Exit Function
                End If
            End If
        Next secondEdge
    Next firstEdge
End Function

Private Function AET_SegmentsCross(ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal byCoordinate As Double, _
                                   ByVal cx As Double, ByVal cy As Double, ByVal dx As Double, ByVal dy As Double) As Boolean
    Dim firstSide As Double
    Dim secondSide As Double
    Dim thirdSide As Double
    Dim fourthSide As Double
    firstSide = (bx - ax) * (cy - ay) - (byCoordinate - ay) * (cx - ax)
    secondSide = (bx - ax) * (dy - ay) - (byCoordinate - ay) * (dx - ax)
    thirdSide = (dx - cx) * (ay - cy) - (dy - cy) * (ax - cx)
    fourthSide = (dx - cx) * (byCoordinate - cy) - (dy - cy) * (bx - cx)
    AET_SegmentsCross = (firstSide * secondSide < 0# And thirdSide * fourthSide < 0#)
End Function
