Imports M4D.Treegram.Core.Entities
Imports M4D.Deixi.Core
Imports Treegram.GeomKernel.BasicModeler
Imports Treegram.GeomKernel.BasicModeler.Surface
Imports SharpDX
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.GeomLibrary
Imports Treegram.GeomFunctions

Public Module CommonGeomFunctions

    ''' <summary>
    ''' Gravity Center of Points with optional weights.
    ''' </summary>
    ''' <param name="pointsList"></param>
    ''' <param name="weightsList"></param>
    ''' <returns></returns>
    Public Function GetBarycenterFromPoints(pointsList As List(Of Vector3), Optional weightsList As List(Of Double) = Nothing) As Vector3
        Dim baryX, baryY, baryZ As Double
        Dim weightsSum As Double
        Dim i As Integer = 0
        For Each myPt In pointsList
            Dim myWeight As Double
            If weightsList Is Nothing Then
                myWeight = 1
            Else
                myWeight = weightsList(i)
            End If
            baryX += myWeight * myPt.X
            baryY += myWeight * myPt.Y
            baryZ += myWeight * myPt.Z
            weightsSum += myWeight
            i += 1
        Next
        baryX /= weightsSum
        baryY /= weightsSum
        baryZ /= weightsSum
        Return New Vector3(baryX, baryY, baryZ)

    End Function

    Public Function BiggestProfile(boundaries As List(Of List(Of Curve2D))) As List(Of Curve2D)
        Dim biggerProfileLength = Double.NegativeInfinity
        Dim biggerIndex As Integer
        Dim j As Integer = 0
        For Each bound In boundaries
            Dim profLength = bound.Select(Function(l) l.ToBasicModelerLine.Length).Sum
            If profLength > biggerProfileLength Then
                biggerProfileLength = profLength
                biggerIndex = j
            End If
            j += 1
        Next
        Return boundaries(biggerIndex)
    End Function

    Public Function ProfileSelfIntersecting(cleanedProfile As Surface) As Boolean
        Dim i, j As Integer
        Dim isSelfIntersectiong = False
        For i = 0 To cleanedProfile.Edges.Count - 2
            Dim firstEdge = cleanedProfile.Edges(i)
            For j = 1 To cleanedProfile.Edges.Count - 1
                Dim secondEdge = cleanedProfile.Edges(j)
                'Same or adjacent indexes
                If Math.Abs(i - j) <= 1 Or Math.Abs(i - j) = cleanedProfile.Edges.Count - 1 Then
                    Continue For
                End If
                'Too far
                If Treegram.GeomKernel.BasicModeler.Distance(firstEdge, secondEdge, False) > 0.001 Then
                    Continue For
                End If
                'If firstEdge.StartPoint = secondEdge.EndPoint Or firstEdge.EndPoint = secondEdge.StartPoint Then
                '    Continue For
                'End If
                Dim er As ErrorCode
                Dim intersPt = Treegram.GeomKernel.BasicModeler.GetIntersection(firstEdge.StartPoint, firstEdge.EndPoint, secondEdge.StartPoint, secondEdge.EndPoint, er, False)
                If er = ErrorCode.CoincidentLines Or intersPt IsNot Nothing Then
                    isSelfIntersectiong = True
                    Exit For
                End If
            Next
            If isSelfIntersectiong Then
                Exit For
            End If
        Next

        Return isSelfIntersectiong
    End Function

    Public Function DoesPtBelongToSurface(surf As Surface, pt As Treegram.GeomKernel.BasicModeler.Point) As Boolean
        For Each innerSurf In surf.InnerSurfaces
            If innerSurf.WindingNumber(pt) > 0 Then
                Return False
            End If
        Next
        If surf.WindingNumber(pt) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function XYmodelsMinMax(models As List(Of Model), Uvec As Vector2) As Tuple(Of Double, Double)

        Dim minU = Double.PositiveInfinity
        Dim maxU = Double.NegativeInfinity

        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim Vvec = Uvec.ToBasicModelerVector.Rotate2(Math.PI / 2)
        Dim Zvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)

        If models Is Nothing OrElse models.Count = 0 Then
            Throw New Exception("No Model for this object ! Did you well load geometry before ?")
        End If
        For Each model In models 'Loop on models
            For Each vertice In model.VerticesTesselation 'Loop on vertices
                Dim pt As New Treegram.GeomKernel.BasicModeler.Point(vertice.X, vertice.Y, 0)
                Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(pt, originPt, Uvec.ToBasicModelerVector, Vvec, Zvec)
                If ptInUV.X < minU Then
                    minU = CDbl(Math.Round(ptInUV.X, 5))
                End If
                If ptInUV.X > maxU Then
                    maxU = CDbl(Math.Round(ptInUV.X, 5))
                End If
            Next
        Next
        Return New Tuple(Of Double, Double)(minU, maxU)
    End Function

    Private Function ConstructBoxFromHorizontalProfile(horizProfile As List(Of Treegram.GeomKernel.BasicModeler.Point), height As Double) As List(Of Treegram.GeomKernel.BasicModeler.Surface)

        Dim zVec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        Dim horizTopProfile As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        For Each pt In horizProfile
            horizTopProfile.Add(pt)
        Next

        'Translate points
        horizTopProfile(0).Translate(zVec, height) 'bottom left (seen from bottom)
        horizTopProfile(1).Translate(zVec, height) 'bottom right
        horizTopProfile(2).Translate(zVec, height) 'top right
        horizTopProfile(3).Translate(zVec, height) 'top left

        'Create surfaces (seen from bottom)
        Dim bottomList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizProfile(0), horizProfile(1), horizTopProfile(1), horizTopProfile(0)}
        Dim topList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizProfile(3), horizProfile(2), horizTopProfile(2), horizTopProfile(3)}
        Dim leftList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizProfile(0), horizProfile(3), horizTopProfile(3), horizTopProfile(0)}
        Dim rightList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizProfile(1), horizProfile(2), horizTopProfile(2), horizTopProfile(1)}
        Dim frontList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizProfile(0), horizProfile(1), horizProfile(2), horizProfile(3)}
        Dim backList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {horizTopProfile(0), horizTopProfile(1), horizTopProfile(2), horizTopProfile(3)}
        Dim surfacesList As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(bottomList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(topList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(leftList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(rightList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(frontList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(backList, SurfaceType.Polygon))
        Return surfacesList
    End Function

    Private Function ConstructBoxFromVerticalProfile(TOTALWIDTHADDED As Double, TOTALTHICKNESSADDED As Double, HEIGHTADDED As Double, wallThickness As Object, frontPtList As List(Of Treegram.GeomKernel.BasicModeler.Point), backPtList As List(Of Treegram.GeomKernel.BasicModeler.Point), openingXvec As Treegram.GeomKernel.BasicModeler.Vector, openingYvec As Treegram.GeomKernel.BasicModeler.Vector, zVec As Treegram.GeomKernel.BasicModeler.Vector) As List(Of Treegram.GeomKernel.BasicModeler.Surface)

        'Translate points
        '- front
        frontPtList(0).Translate(openingXvec.Reverse, TOTALWIDTHADDED / 2) 'bottom left
        frontPtList(0).Translate(openingYvec.Reverse, TOTALTHICKNESSADDED / 2)

        frontPtList(1).Translate(openingXvec, TOTALWIDTHADDED / 2) 'bottom right
        frontPtList(1).Translate(openingYvec.Reverse, TOTALTHICKNESSADDED / 2)

        frontPtList(2).Translate(openingXvec, TOTALWIDTHADDED / 2) 'top right
        frontPtList(2).Translate(openingYvec.Reverse, TOTALTHICKNESSADDED / 2)
        frontPtList(2).Translate(zVec, HEIGHTADDED)

        frontPtList(3).Translate(openingXvec.Reverse, TOTALWIDTHADDED / 2) 'top left
        frontPtList(3).Translate(openingYvec.Reverse, TOTALTHICKNESSADDED / 2)
        frontPtList(3).Translate(zVec, HEIGHTADDED)

        '- back
        backPtList(0).Translate(openingXvec.Reverse, TOTALWIDTHADDED / 2) 'bottom left
        backPtList(0).Translate(openingYvec, wallThickness + TOTALTHICKNESSADDED / 2)

        backPtList(1).Translate(openingXvec, TOTALWIDTHADDED / 2) 'bottom right
        backPtList(1).Translate(openingYvec, wallThickness + TOTALTHICKNESSADDED / 2)

        backPtList(2).Translate(openingXvec, TOTALWIDTHADDED / 2) 'top right
        backPtList(2).Translate(openingYvec, wallThickness + TOTALTHICKNESSADDED / 2)
        backPtList(2).Translate(zVec, HEIGHTADDED)

        backPtList(3).Translate(openingXvec.Reverse, TOTALWIDTHADDED / 2) 'top left
        backPtList(3).Translate(openingYvec, wallThickness + TOTALTHICKNESSADDED / 2)
        backPtList(3).Translate(zVec, HEIGHTADDED)

        'Create surfaces
        Dim bottomList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(1), backPtList(1), backPtList(0)}
        Dim topList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(3), frontPtList(2), backPtList(2), backPtList(3)}
        Dim leftList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(3), backPtList(3), backPtList(0)}
        Dim rightList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(1), frontPtList(2), backPtList(2), backPtList(1)}
        Dim frontList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {frontPtList(0), frontPtList(1), frontPtList(2), frontPtList(3)}
        Dim backList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {backPtList(0), backPtList(1), backPtList(2), backPtList(3)}
        Dim surfacesList As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(bottomList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(topList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(leftList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(rightList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(frontList, SurfaceType.Polygon))
        surfacesList.Add(New Treegram.GeomKernel.BasicModeler.Surface(backList, SurfaceType.Polygon))
        Return surfacesList
    End Function


    Public Function GetSuperpositionArea(surfA As Treegram.GeomKernel.BasicModeler.Surface, surfB As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Dim superpositionArea As Double = 0.0
        If surfA.GetNormal.X <> 0 Or surfA.GetNormal.Y <> 0 Or surfB.GetNormal.X <> 0 Or surfB.GetNormal.Y <> 0 OrElse surfA.Points(0).Z <> surfB.Points(0).Z Then
            Throw New Exception("Surfaces must belong to the same horizontal plane")
        End If
        If Treegram.GeomKernel.BasicModeler.Distance(surfA, surfB) = 0 Then 'Voir si cette méthode ne ralentit pas le process
            Dim intersSurf As New List(Of Surface)
            Try 'A AMÉLIORER - Mais après quelques vérifs ce sont les Contacts qui merdent et non les clashs !
                intersSurf = Treegram.GeomKernel.BasicModeler.SurfacePolyBoolIntersection(surfA, surfB)
            Catch ex As Exception
            End Try
            intersSurf = SurfaceListCleaningMethods(intersSurf)
            If intersSurf.Count > 0 Then
                superpositionArea = intersSurf.Select(Function(S) S.Area).Sum
            End If

            ''DEBUG - Using Rhino....
            'Dim debug As Boolean = False
            'If debug AndAlso (superpositionArea / surfA.Area < 0.99 Or superpositionArea / surfB.Area < 0.99) Then
            '    Dim nullDistanceList As New List(of Treegram.GeomKernel.BasicModeler.Surface) From {surfA, surfB}
            '    DrawPolylines(nullDistanceList, "SurfToSuperpose", True, System.Drawing.Color.DarkOrange)
            '    DrawPolylines(intersSurf, "Superposition", True, System.Drawing.Color.MediumPurple)
            '    MessageBox.Show("Click OK to continue")
            'End If
        End If
        Return superpositionArea
    End Function

    Public Function GetTheBiggestBoundary(boundaries As List(Of List(Of Line))) As List(Of Line)
        Dim boundary As List(Of Treegram.GeomKernel.BasicModeler.Line)
        Dim surfMax = Double.NegativeInfinity
        For Each bound In boundaries 'Keep the biggest one
            If bound.Count < 3 Then
                Continue For
            End If
            Dim surf As New Treegram.GeomKernel.BasicModeler.Surface(bound.Select(Function(w) w.StartPoint).ToList(), SurfaceType.Polygon)
            If Math.Round(surf.Area, 3) > surfMax Then
                surfMax = Math.Round(surf.Area, 3)
                boundary = bound
            End If
        Next
        If surfMax = Double.NegativeInfinity Then
            Throw New Exception("Impossible")
        End If

        Return boundary
    End Function

    Public Function SurfaceListCleaningMethods(intersSurfList As Object) As List(Of Surface)
        Dim cleanedIntersSurf As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        For Each surf In intersSurfList
            Dim cleanedSurf = SurfaceCleaningMethods(surf, 4)
            If cleanedSurf IsNot Nothing Then
                cleanedIntersSurf.Add(cleanedSurf)
            End If
        Next

        Return cleanedIntersSurf
    End Function

    Public Function SurfaceCleaningMethods(contourSurface As Surface, Optional mathRound As Integer = 4) As Surface
        Dim firstSurface As New Surface(contourSurface.Points, SurfaceType.Polygon)
        If contourSurface.GetNormal Is Nothing Then
            Return Nothing
        End If
        Dim startArea = contourSurface.Area
        Dim cleaningPtDone = False
        Do While Not cleaningPtDone
            Dim lastPtCount = contourSurface.Points.Count
            contourSurface = contourSurface.CleanPoints(mathRound)
            If contourSurface Is Nothing OrElse contourSurface.Points.Count = lastPtCount Then
                cleaningPtDone = True
            End If
        Loop
        If contourSurface Is Nothing Then
            Return Nothing
            Throw New Exception("Cleaning method failed")
        End If

        Dim cleaningEdgeDone = False
        Do While Not cleaningEdgeDone
            Dim lastPtCount = contourSurface.Points.Count
            contourSurface = contourSurface.CleanEdges(3, Math.Pow(10, -mathRound))
            If contourSurface Is Nothing OrElse contourSurface.Points.Count = lastPtCount Then
                cleaningEdgeDone = True
            End If
        Loop
        If contourSurface Is Nothing Then
            Return Nothing
            Throw New Exception("Cleaning method failed")
        End If
        'If Math.Abs(startArea - contourSurface.Area) > 10 ^ (-2 * (mathRound - 1)) Then
        If Math.Abs(startArea - contourSurface.Area) > 0.01 Then 'Ne pas réduire car la précision de certaines surfaces est de l'ordre de 10-3
            Return firstSurface
            Throw New Exception("Cleaning method failed")
        End If
        Return contourSurface
    End Function

    Public Function SetElevationToSurface(mySurfList As List(Of Treegram.GeomKernel.BasicModeler.Surface), elevation As Double) As List(Of Treegram.GeomKernel.BasicModeler.Surface)
        Dim newSurfList As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        For Each mySurf In mySurfList
            Dim myPtList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
            Dim newInnerBoundList As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
            For Each pt In mySurf.Points
                myPtList.Add(New Treegram.GeomKernel.BasicModeler.Point(pt.X, pt.Y, elevation)) 'Add elevation to the surface
            Next
            For Each bound In mySurf.InnerSurfaces
                Dim newBound As New List(Of Treegram.GeomKernel.BasicModeler.Point)
                For Each pt In bound.Points
                    newBound.Add(New Treegram.GeomKernel.BasicModeler.Point(pt.X, pt.Y, elevation)) 'Add elevation to the surface
                Next
                newInnerBoundList.Add(New Treegram.GeomKernel.BasicModeler.Surface(newBound, SurfaceType.Polygon))
            Next
            Dim newSurf As New Treegram.GeomKernel.BasicModeler.Surface(myPtList, SurfaceType.Polygon) With {.InnerSurfaces = newInnerBoundList}
            newSurfList.Add(newSurf)
        Next
        Return newSurfList
    End Function

    Public Function Get3DFromHorizontalProfile(overlapProfile As Surface, elevInf As Double, elevSup As Double) As List(Of Surface)
        Dim bottomPtsList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        Dim topPtsList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        Dim allSurfacesList As New List(Of Surface)
        For Each edge As Line In overlapProfile.Edges
            'Bottom
            Dim startBotPt = edge.StartPoint.Clone
            startBotPt.Z = elevInf
            Dim endBotPt = edge.EndPoint.Clone
            endBotPt.Z = elevInf
            bottomPtsList.Add(startBotPt)
            'Top
            Dim startTopPt = edge.StartPoint.Clone
            startTopPt.Z = elevSup
            Dim endTopPt = edge.EndPoint.Clone
            endTopPt.Z = elevSup
            topPtsList.Add(startTopPt)
            'Side Surface
            Dim sidePtsList As New List(Of Treegram.GeomKernel.BasicModeler.Point) From {startBotPt, startTopPt, endTopPt, endBotPt}
            Dim sideSurf As New Surface(sidePtsList, Surface.SurfaceType.Polygon)
            allSurfacesList.Add(sideSurf)
        Next
        Dim botSurface As New Surface(bottomPtsList, Surface.SurfaceType.Polygon)
        allSurfacesList.Add(botSurface)
        Dim topSurface As New Surface(topPtsList, Surface.SurfaceType.Polygon)
        allSurfacesList.Add(topSurface)
        Return allSurfacesList
    End Function



    ''' <summary></summary>
    ''' <param name="vecU"></param>
    ''' <param name="vecV"></param>
    ''' <param name="ptList"></param>
    ''' <returns>MinU, MaxU, MinV, MaxV</returns>
    Private Function ProfileExtremums(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, ptList As List(Of Treegram.GeomKernel.BasicModeler.Point)) As Tuple(Of Double, Double, Double, Double)
        Dim _MinU = Double.PositiveInfinity
        Dim _MaxU = Double.NegativeInfinity
        Dim _MinV = Double.PositiveInfinity
        Dim _MaxV = Double.NegativeInfinity
        Dim _MinZ = Double.PositiveInfinity
        Dim _MaxZ = Double.NegativeInfinity

        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        For Each pt In ptList

            Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(pt, originPt, vecU, vecV, axisZvec)

            If ptInUV.X < _MinU Then
                _MinU = CDbl(Math.Round(ptInUV.X, 3))
            End If
            If ptInUV.X > _MaxU Then
                _MaxU = CDbl(Math.Round(ptInUV.X, 3))
            End If
            If ptInUV.Y < _MinV Then
                _MinV = CDbl(Math.Round(ptInUV.Y, 3))
            End If
            If ptInUV.Y > _MaxV Then
                _MaxV = CDbl(Math.Round(ptInUV.Y, 3))
            End If
        Next
        Return New Tuple(Of Double, Double, Double, Double)(_MinU, _MaxU, _MinV, _MaxV)
    End Function

    Public Function LineMinU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, line As Treegram.GeomKernel.BasicModeler.Line) As Double
        Return ProfileExtremums(vecU, vecV, {line.StartPoint, line.EndPoint}.ToList).Item1
    End Function
    Public Function LineMaxU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, line As Treegram.GeomKernel.BasicModeler.Line) As Double
        Return ProfileExtremums(vecU, vecV, {line.StartPoint, line.EndPoint}.ToList).Item2
    End Function
    Public Function LineMinV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, line As Treegram.GeomKernel.BasicModeler.Line) As Double
        Return ProfileExtremums(vecU, vecV, {line.StartPoint, line.EndPoint}.ToList).Item3
    End Function
    Public Function LineMaxV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, line As Treegram.GeomKernel.BasicModeler.Line) As Double
        Return ProfileExtremums(vecU, vecV, {line.StartPoint, line.EndPoint}.ToList).Item4
    End Function

    Public Function ProfileMinU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, surface As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Return ProfileExtremums(vecU, vecV, surface.Points).Item1
    End Function
    Public Function ProfileMaxU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, surface As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Return ProfileExtremums(vecU, vecV, surface.Points).Item2
    End Function
    Public Function ProfileMinV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, surface As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Return ProfileExtremums(vecU, vecV, surface.Points).Item3
    End Function
    Public Function ProfileMaxV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, surface As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Return ProfileExtremums(vecU, vecV, surface.Points).Item4
    End Function

    Public Function ProfileMinZ(profile As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Dim _MinZ = Double.PositiveInfinity
        For Each pt In profile.Points
            If pt.Z < _MinZ Then
                _MinZ = CDbl(Math.Round(pt.Z, 3))
            End If
        Next
        Return _MinZ
    End Function
    Public Function ProfileMaxZ(profile As Treegram.GeomKernel.BasicModeler.Surface) As Double
        Dim _MaxZ = Double.NegativeInfinity
        For Each pt In profile.Points
            If pt.Z > _MaxZ Then
                _MaxZ = CDbl(Math.Round(pt.Z, 3))
            End If
        Next
        Return _MaxZ
    End Function


    <Obsolete>
    Public Function ProfileMaxU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MaxU = Double.NegativeInfinity
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        For Each line In profile
            Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(line.StartPoint, originPt, vecU, vecV, axisZvec)
            If ptInUV.X > _MaxU Then
                _MaxU = CDbl(Math.Round(ptInUV.X, 3))
            End If
        Next
        Return _MaxU
    End Function
    <Obsolete>
    Public Function ProfileMaxV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MaxV = Double.NegativeInfinity
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        For Each line In profile
            Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(line.StartPoint, originPt, vecU, vecV, axisZvec)
            If ptInUV.Y > _MaxV Then
                _MaxV = CDbl(Math.Round(ptInUV.Y, 3))
            End If
        Next
        Return _MaxV
    End Function
    <Obsolete>
    Public Function ProfileMinU(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MinU = Double.PositiveInfinity
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        For Each line In profile
            Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(line.StartPoint, originPt, vecU, vecV, axisZvec)
            If ptInUV.X < _MinU Then
                _MinU = CDbl(Math.Round(ptInUV.X, 3))
            End If
        Next
        Return _MinU
    End Function
    <Obsolete>
    Public Function ProfileMinV(vecU As Treegram.GeomKernel.BasicModeler.Vector, vecV As Treegram.GeomKernel.BasicModeler.Vector, profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MinV = Double.PositiveInfinity
        Dim originPt As New Treegram.GeomKernel.BasicModeler.Point(0, 0, 0)
        Dim axisZvec As New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1)
        For Each line In profile
            Dim ptInUV = Treegram.GeomKernel.BasicModeler.ChangeCoordPointInAnotherAxisSyst(line.StartPoint, originPt, vecU, vecV, axisZvec)
            If ptInUV.Y < _MinV Then
                _MinV = CDbl(Math.Round(ptInUV.Y, 3))
            End If
        Next
        Return _MinV

    End Function
    <Obsolete>
    Public Function ProfileMinZ(profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MinZ = Double.PositiveInfinity
        For Each line In profile
            If line.StartPoint.Z < _MinZ Then
                _MinZ = CDbl(Math.Round(line.StartPoint.Z, 3))
            End If
        Next
        Return _MinZ
    End Function
    <Obsolete>
    Public Function ProfileMaxZ(profile As List(Of Treegram.GeomKernel.BasicModeler.Line)) As Double
        Dim _MaxZ = Double.NegativeInfinity
        For Each line In profile
            If line.StartPoint.Z > _MaxZ Then
                _MaxZ = CDbl(Math.Round(line.StartPoint.Z, 3))
            End If
        Next
        Return _MaxZ
    End Function
End Module

Public Module Intersections2D

    Public Function GetIntersectionFacingDownward(objTgm As MetaObject, elevation As Double) As List(Of List(Of Treegram.GeomKernel.BasicModeler.Line))
        Dim boundaries As New List(Of List(Of Treegram.GeomKernel.BasicModeler.Line))
        Dim surfaces As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        For Each model In Treegram.GeomFunctions.Models.GetGeometryModels(objTgm)
            Dim sectionPlane As New Treegram.GeomKernel.BasicModeler.Plane(New Treegram.GeomKernel.BasicModeler.Point(0, 0, elevation), New Treegram.GeomKernel.BasicModeler.Vector(0, 0, 1))
            Dim projsSurf = Treegram.GeomFunctions.Models.GetSectionSurfaces(model, sectionPlane)

            For Each projSurf In projsSurf
                'pour se débarasser des doublons - A AMELIORER - sert-il vraiment ???
                For Each surf In surfaces
                    If projSurf = surf Then
                        GoTo nextSurf
                    End If
                Next
                surfaces.Add(projSurf)

                Dim newEdges As New List(Of Treegram.GeomKernel.BasicModeler.Line)
                Dim i As Integer
                For i = projSurf.Edges.Count - 1 To 0 Step -1
                    Dim stPt = projSurf.Edges(i).StartPoint
                    Dim endPt = projSurf.Edges(i).EndPoint
                    newEdges.Add(New Treegram.GeomKernel.BasicModeler.Line(New Treegram.GeomKernel.BasicModeler.Point(endPt.X, endPt.Y), New Treegram.GeomKernel.BasicModeler.Point(stPt.X, stPt.Y)))
                Next
                'For i = 0 To projSurf.Edges.Count - 1
                '    newEdges.Add(new Treegram.GeomKernel.BasicModeler.Line(projSurf.Edges(i).StartPoint, projSurf.Edges(i).EndPoint))
                'Next
                boundaries.Add(newEdges)
nextSurf:
            Next
        Next
        Return boundaries
    End Function

    Public Function GetTopFaceFacingDownward(objTgm As MetaObject, elevation As Double) As List(Of List(Of Treegram.GeomKernel.BasicModeler.Line))
        'Get ALL faces
        Dim faces As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        For Each model In Treegram.GeomFunctions.Models.GetGeometryModels(objTgm)
            Dim subFaces = Treegram.GeomFunctions.Models.GetModelAsFaces(model)
            For Each subFace In subFaces
                faces.Add(subFace)
            Next
        Next

        'Compare average elevation
        Dim highestDb = Double.NegativeInfinity
        Dim highestSurf As Treegram.GeomKernel.BasicModeler.Surface
        For Each surf In faces
            Dim averageHeight = surf.Points.Select(Function(pt) pt.Z).Average
            If averageHeight > highestDb Then
                highestDb = averageHeight
                highestSurf = surf
            End If
        Next

        ''Set right level
        Dim correctElevPtList As New List(Of Treegram.GeomKernel.BasicModeler.Point)
        For Each pt In highestSurf.Points
            correctElevPtList.Add(New Treegram.GeomKernel.BasicModeler.Point(pt.X, pt.Y, elevation))
        Next
        Dim correctElevSurf As New Treegram.GeomKernel.BasicModeler.Surface(correctElevPtList, SurfaceType.Polygon)

        'Set right orientation
        Dim newEdges As New List(Of Treegram.GeomKernel.BasicModeler.Line)
        Dim i As Integer
        For i = correctElevSurf.Edges.Count - 1 To 0 Step -1
            newEdges.Add(New Treegram.GeomKernel.BasicModeler.Line(correctElevSurf.Edges(i).EndPoint, correctElevSurf.Edges(i).StartPoint))
        Next

        Return New List(Of List(Of Treegram.GeomKernel.BasicModeler.Line)) From {newEdges}
    End Function

End Module

Public Module Projections2D

    Private Function GetTopFaceFacingDownward(objTgm As MetaObject) As List(Of List(Of Treegram.GeomKernel.BasicModeler.Line))
        'Get ALL faces
        Dim faces As New List(Of Treegram.GeomKernel.BasicModeler.Surface)
        For Each model In Treegram.GeomFunctions.Models.GetGeometryModels(objTgm)
            Dim subFaces = Treegram.GeomFunctions.Models.GetModelAsFaces(model)
            For Each subFace In subFaces
                faces.Add(subFace)
            Next
        Next

        'Compare average elevation
        Dim highestDb = Double.NegativeInfinity
        Dim highestSurf As Treegram.GeomKernel.BasicModeler.Surface
        For Each surf In faces
            Dim averageHeight = surf.Points.Select(Function(pt) pt.Z).Average
            If averageHeight > highestDb Then
                highestDb = averageHeight
                highestSurf = surf
            End If
        Next

        'Set right orientation
        Dim newEdges As New List(Of Treegram.GeomKernel.BasicModeler.Line)
        Dim i As Integer
        For i = highestSurf.Edges.Count - 1 To 0 Step -1
            Dim stPt = highestSurf.Edges(i).StartPoint
            Dim endPt = highestSurf.Edges(i).EndPoint
            newEdges.Add(New Treegram.GeomKernel.BasicModeler.Line(New Treegram.GeomKernel.BasicModeler.Point(endPt.X, endPt.Y), New Treegram.GeomKernel.BasicModeler.Point(stPt.X, stPt.Y)))
        Next

        Return New List(Of List(Of Treegram.GeomKernel.BasicModeler.Line)) From {newEdges}
    End Function

    Public Function GetHorizProjFacingDownward(objTgm As MetaObject, Optional mathRound As Integer = 4) As List(Of List(Of Curve2D))
        Dim boundaries As New List(Of List(Of Curve2D))

        Dim objModels
        If objTgm.GetTgmType = "IfcWall" Then
            objModels = Treegram.GeomFunctions.Models.GetGeometryModels(objTgm, True) 'Could be gross model for better results !
        ElseIf objTgm.GetTgmType = "IfcCurtainWall" Then
            objModels = GetCurtainWallChildrenModels(objTgm)
        Else
            objModels = Treegram.GeomFunctions.Models.GetGeometryModels(objTgm, True)
        End If

        For Each model In objModels
            Dim projsSurf
            Try
                projsSurf = Treegram.GeomFunctions.Models.Get3dProjectionOnPlane(model, Nothing, mathRound)
            Catch ex As Exception
                Continue For
            End Try
            For Each projSurf As Treegram.GeomKernel.BasicModeler.Surface In projsSurf
                If projSurf.GetNormal Is Nothing Then
                    Continue For
                End If
                Dim newEdges As New List(Of Curve2D)
                Dim i As Integer
                For i = projSurf.Edges.Count - 1 To 0 Step -1
                    Dim newStartPt As New Vector2(projSurf.Edges(i).EndPoint.X, projSurf.Edges(i).EndPoint.Y)
                    Dim newEndPt As New Vector2(projSurf.Edges(i).StartPoint.X, projSurf.Edges(i).StartPoint.Y)
                    newEdges.Add(New Curve2D({newStartPt, newEndPt}.ToList))
                Next
                'Dim ptList As New List(Of Vector2)
                'Dim i As Integer
                'For i = projSurf.Edges.Count - 1 To 0 Step -1
                '    ptList.Add(New Vector2(projSurf.Edges(i).EndPoint.X, projSurf.Edges(i).EndPoint.Y))
                'Next
                boundaries.Add(newEdges)
            Next
        Next
        Return boundaries
    End Function

    Private Function GetCurtainWallChildrenModels(objTgm As MetaObject) As List(Of Model)
        Dim objModels As New List(Of Model)
        For Each childTgm In objTgm.GetChildren(Nothing, Nothing, RelationType.Decomposition).ToList
            If childTgm.GetTgmType <> "IfcWindow" And childTgm.GetTgmType <> "IfcDoor" Then 'Utiliser les CategoryByTgm plutôt...
                objModels.AddRange(Treegram.GeomFunctions.Models.GetGeometryModels(childTgm, True))
            End If
        Next
        Return objModels
    End Function
End Module



