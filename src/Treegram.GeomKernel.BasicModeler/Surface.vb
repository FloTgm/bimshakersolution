

''' <summary>
''' Represents a 2D-Surface
''' </summary>
Public Class Surface

    ''' <summary>
    ''' Types of Surface
    ''' </summary>
    ''' Prefer Polygon to Rectangle
    Public Enum SurfaceType
        Circle
        Rectangle
        Polygon
        Complex
        Ellipse
    End Enum

    ''' <summary>
    ''' List of Points of the Surface. External Boundary. 
    ''' </summary>
    ''' <returns></returns>
    Public Property Points As List(Of Point)

    ''' <summary>
    ''' Each List of Points represent an internal boundary. A hole in the face.
    ''' Be careful. This property is only used to create 3D in Treegram.
    ''' The methods in this class don't take into account innerbounds yet.
    ''' </summary>
    ''' <returns></returns>
    <Obsolete("Take InnerSurfaces")>
    Public Property InnerBounds As List(Of List(Of Point))

    ''' <summary>
    ''' These inner bounds are hole in the main surface.
    ''' </summary>
    ''' <returns></returns>
    Public Property InnerSurfaces As List(Of Surface)

    ''' <summary>
    ''' Type of the Surface
    ''' </summary>
    ''' <returns></returns>
    Public Property Type As SurfaceType


    ''' <summary>
    ''' List of Edges of the Surface
    ''' </summary>
    ''' <returns></returns>
    Public Property Edges As List(Of Curve)

    ''' <summary>
    ''' Creates new Surface by a list of Points
    ''' </summary>
    ''' <param name="points"> List of Points </param>
    ''' <param name="type"> Type of Surface </param>
    Public Sub New(points As List(Of Point), type As SurfaceType)
        Me.Points = points
        Me.Type = type
        Me.InnerSurfaces = New List(Of Surface)
        If type = SurfaceType.Ellipse Then
            If Math.Abs(New Vector(Me.Points(0), Me.Points(1)) * New Vector(Me.Points(0), Me.Points(2))) > Math.Pow(10, -4) Then
                Throw New Exception("Ellipse Axis not perpendicular")
            End If
        End If
        If type = SurfaceType.Rectangle OrElse type = SurfaceType.Polygon Then
            Dim list As New List(Of Curve)
            For k = 0 To points.Count - 2
                list.Add(New Line(points(k), points(k + 1)))
            Next
            list.Add(New Line(points.Last, points.First))
            Me.Edges = list
        End If
    End Sub

    ''' <summary>
    ''' Creates a new Surface by a list of Edges
    ''' </summary>
    ''' <param name="edges"> List od Edges </param>
    ''' <param name="type"> Type of Surface </param>
    Public Sub New(edges As List(Of Curve), type As SurfaceType)
        Me.Edges = edges
        Me.Type = type
        Me.InnerSurfaces = New List(Of Surface)
        Dim list As New List(Of Point)
        For Each edge In edges
            list.Add(edge.StartPoint)
        Next
        Me.Points = list
    End Sub

    ''' <summary>
    ''' Get area of the Surface
    ''' </summary>
    ''' <returns> Area </returns>
    Public ReadOnly Property Area As Double
        Get
            'Dim areaHoles = 0
            'If Me.InnerSurfaces.Count > 0 Then
            '    For Each surface In InnerSurfaces
            '        areaHoles += surface.Area
            '    Next
            'End If
            'Select Case Type
            '    Case SurfaceType.Circle
            '        Dim rayon As Double = Distance(Points(0), Points(1))
            '        Return Math.Round(RoundToSignificantDigits(rayon * rayon * Math.PI, RealPrecision), -DefaultTolerance)
            '    Case SurfaceType.Rectangle
            '        Dim longueur As Double = Distance(Points(0), Points(1))
            '        Dim largeur As Double = Distance(Points(1), Points(2))
            '        Return Math.Round(RoundToSignificantDigits(longueur * largeur, RealPrecision), -DefaultTolerance)
            '    Case SurfaceType.Polygon
            '        Dim result As Double = (Points.Last.X * Points.First.Y - Points.First.X * Points.Last.Y) / 2
            '        Dim r = 1
            '        For j = 0 To Points.Count - 2
            '            result += (Points(j).X * Points(j + 1).Y - Points(j + 1).X * Points(j).Y) / 2
            '            r += 1
            '            If r = 4 Then
            '                result = RoundToSignificantDigits(result, RealPrecision)
            '                r = 0
            '            End If
            '        Next
            '        If r <> 0 Then
            '            result = RoundToSignificantDigits(result, RealPrecision)
            '        End If
            '        Return Math.Abs(Math.Round(result, -DefaultTolerance))
            '    Case SurfaceType.Ellipse
            '        Dim rayonA = Distance(Points(0), Points(1))
            '        Dim rayonB = Distance(Points(0), Points(2))
            '        Return Math.Round(RoundToSignificantDigits(rayonA * rayonB * Math.PI, RealPrecision), -DefaultTolerance)
            '    Case SurfaceType.Complex
            '        Throw New Exception("Not yet implemented")
            '    Case Else
            '        Throw New Exception("Define surface type before checking area")
            'End Select
            Return Math.Round(Me.InternArea, -DefaultTolerance)
        End Get
    End Property

    ''' <summary>
    ''' Get Area of 3D Surface.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Area3D As Double
        Get
            'Dim areaHoles = 0
            'If Me.InnerSurfaces.Count > 0 Then
            '    For Each surface In InnerSurfaces
            '        areaHoles += surface.Area3D
            '    Next
            'End If
            'Select Case Type
            '    Case SurfaceType.Polygon
            '        Dim result As Double = Vector.CrossProduct(New Vector(Points.First, Points(1)), New Vector(Points.First, Points(2))).InternLength / 2
            '        Dim r = 1
            '        For j = 2 To Points.Count - 2
            '            result += Vector.CrossProduct(New Vector(Points.First, Points(j)), New Vector(Points.First, Points(j + 1))).InternLength / 2
            '            r += 1
            '            If r = 4 Then
            '                result = RoundToSignificantDigits(result, RealPrecision)
            '                r = 0
            '            End If
            '        Next
            '        If r <> 0 Then
            '            result = RoundToSignificantDigits(result, RealPrecision)
            '        End If
            '        Return Math.Abs(Math.Round(result, -DefaultTolerance))
            '    Case Else
            '        Return Area
            'End Select
            Return Math.Round(Me.InternArea3D, -DefaultTolerance)
        End Get
    End Property

    Friend ReadOnly Property InternArea As Double
        Get
            Dim areaHoles = 0
            If Me.InnerSurfaces.Count > 0 Then
                For Each surface In InnerSurfaces
                    areaHoles += surface.InternArea
                Next
            End If
            Select Case Type
                Case SurfaceType.Circle
                    Dim rayon As Double = Distance(Points(0), Points(1))
                    Return RoundToSignificantDigits(rayon * rayon * Math.PI - areaHoles, RealPrecision)
                Case SurfaceType.Rectangle
                    Dim longueur As Double = Distance(Points(0), Points(1))
                    Dim largeur As Double = Distance(Points(1), Points(2))
                    Return RoundToSignificantDigits(longueur * largeur - areaHoles, RealPrecision)
                Case SurfaceType.Polygon
                    Dim result As Double = (Points.Last.X * Points.First.Y - Points.First.X * Points.Last.Y) / 2
                    Dim r = 1
                    For j = 0 To Points.Count - 2
                        result += (Points(j).X * Points(j + 1).Y - Points(j + 1).X * Points(j).Y) / 2
                        r += 1
                        If r = 4 Then
                            result = RoundToSignificantDigits(result, RealPrecision)
                            r = 0
                        End If
                    Next
                    If r <> 0 Then
                        result = RoundToSignificantDigits(result, RealPrecision)
                    End If
                    Return RoundToSignificantDigits(Math.Abs(result) - areaHoles, RealPrecision)
                Case SurfaceType.Ellipse
                    Dim rayonA = Distance(Points(0), Points(1))
                    Dim rayonB = Distance(Points(0), Points(2))
                    Return RoundToSignificantDigits(rayonA * rayonB * Math.PI - areaHoles, RealPrecision)
                Case SurfaceType.Complex
                    Throw New Exception("Not yet implemented")
                Case Else
                    Throw New Exception("Define surface type before checking area")
            End Select
        End Get
    End Property

    Friend ReadOnly Property InternArea3D As Double
        Get
            Select Case Type
                Case SurfaceType.Polygon
                    Dim areaHoles = 0
                    If Me.InnerSurfaces.Count > 0 Then
                        For Each surface In InnerSurfaces
                            areaHoles += surface.InternArea3D
                        Next
                    End If
                    Dim result As Double = Vector.CrossProduct(New Vector(Points.First, Points(1)), New Vector(Points.First, Points(2))).InternLength / 2
                    Dim r = 1
                    For j = 2 To Points.Count - 2
                        result += Vector.CrossProduct(New Vector(Points.First, Points(j)), New Vector(Points.First, Points(j + 1))).InternLength / 2
                        r += 1
                        If r = 4 Then
                            result = RoundToSignificantDigits(result, RealPrecision)
                            r = 0
                        End If
                    Next
                    If r <> 0 Then
                        result = RoundToSignificantDigits(result, RealPrecision)
                    End If
                    Return RoundToSignificantDigits(Math.Abs(result) - areaHoles, RealPrecision)
                Case Else
                    Return InternArea3D
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Used by GetInnerAngles.
    ''' </summary>
    Public SensDeLecture As Boolean? = Nothing

    Public Shared Operator +(surface1 As Surface, surface2 As Surface) As Surface
        Throw New NotImplementedException
    End Operator

    Public Shared Operator =(surf1 As Surface, surf2 As Surface) As Boolean
        If surf1 Is Nothing And surf2 Is Nothing Then
            Return True
        ElseIf surf1 Is Nothing And surf2 IsNot Nothing Or surf1 IsNot Nothing And surf2 Is Nothing Then
            Return False
        Else
            If surf1.Type = surf2.Type Then
                If surf1.Type = SurfaceType.Circle Then
                    If surf1.Points(0) = surf2.Points(0) Then
                        Dim rayon1 As Double = Distance(surf1.Points(0), surf2.Points(1))
                        Dim rayon2 As Double = Distance(surf2.Points(0), surf2.Points(1))
                        If rayon1 = rayon2 Then
                            Return True
                        End If
                    End If
                ElseIf surf1.Type = SurfaceType.Ellipse Then
                    If surf1.Points(0) = surf2.Points(0) Then
                        Dim rayonA1 As Double = Distance(surf1.Points(0), surf2.Points(1))
                        Dim rayonA2 As Double = Distance(surf2.Points(0), surf2.Points(1))
                        Dim rayonB1 As Double = Distance(surf1.Points(0), surf2.Points(2))
                        Dim rayonB2 As Double = Distance(surf2.Points(0), surf2.Points(2))
                        If (rayonA1 = rayonA2 AndAlso rayonB1 = rayonB2) OrElse (rayonA1 = rayonB2 AndAlso rayonA2 = rayonB1) Then
                            Return True
                        End If
                    End If
                Else
                    If surf1.Points.Count = surf2.Points.Count Then
                        Dim vertices1 As New List(Of VertexData) From {
                        New VertexData(surf1.Points.First)
                                        }
                        For k = 1 To surf1.Points.Count - 1
                            vertices1.Add(New VertexData(surf1.Points(k), vertices1(k - 1)))
                            vertices1(k - 1).Suiv = vertices1(k)
                        Next
                        vertices1.First.Prec = vertices1.Last
                        vertices1.Last.Suiv = vertices1.First

                        Dim vertices2 As New List(Of VertexData) From {
                        New VertexData(surf2.Points.First)
                                            }
                        For k = 1 To surf2.Points.Count - 1
                            vertices2.Add(New VertexData(surf2.Points(k), vertices2(k - 1)))
                            vertices2(k - 1).Suiv = vertices2(k)
                        Next
                        vertices2.First.Prec = vertices2.Last
                        vertices2.Last.Suiv = vertices2.First

                        Dim rank = -1
                        For Each vertex In vertices2
                            If vertices1.First = vertex Then
                                rank = vertices2.IndexOf(vertex)
                                Exit For
                            End If
                        Next
                        If Not rank = -1 Then
                            Dim current1 = vertices1.First
                            Dim current2 = vertices2(rank)
                            Dim sens As Boolean = Nothing
                            If current1.Suiv = current2.Suiv Then
                                sens = True
                            End If
                            If current1.Suiv = current2.Prec Then
                                sens = False
                            End If
                            If Not IsNothing(sens) Then
                                For k = 1 To vertices1.Count - 1
                                    current1 = current1.Suiv
                                    If sens Then
                                        current2 = current2.Suiv
                                    Else
                                        current2 = current2.Prec
                                    End If
                                    If current1 <> current2 Then
                                        Return False
                                        Exit For
                                    End If
                                Next
                                Return True
                            End If
                        End If
                    End If
                End If
            End If
        End If

        Return False
    End Operator

    Public Shared Operator <>(surf1 As Surface, surf2 As Surface) As Boolean
        Return Not surf1 = surf2
    End Operator

    ''' <summary>
    ''' Get the list of intern angles of polygons and the reading direction of the points
    ''' </summary>
    ''' <returns></returns>
    Public Function GetInnerAngles() As Tuple(Of List(Of Double), Boolean)
        If Type = SurfaceType.Rectangle OrElse Type = SurfaceType.Polygon Then
            Dim sensAAngles As New List(Of Double), sensBAngles As New List(Of Double)
            Dim sommeAnglesSensA As Double = 0.0, sommeAnglesSensB As Double = 0.0

            For i = 1 To Points.Count
                Dim sensAPA, sensAPB, sensAPC As Point
                If i = 1 Then
                    sensAPA = Points.Last
                    sensAPB = Points(i - 1)
                    sensAPC = Points(i)
                ElseIf i = Points.Count Then
                    sensAPA = Points(i - 2)
                    sensAPB = Points(i - 1)
                    sensAPC = Points.First
                Else
                    sensAPA = Points(i - 2)
                    sensAPB = Points(i - 1)
                    sensAPC = Points(i)
                End If

                Dim vsAPAb = New Vector() With {.X = sensAPA.X - sensAPB.X, .Y = sensAPA.Y - sensAPB.Y}
                Dim vsAPBc = New Vector() With {.X = sensAPC.X - sensAPB.X, .Y = sensAPC.Y - sensAPB.Y}

                Dim angleSensAPAb = Math.Atan2(vsAPAb.ProperY, vsAPAb.ProperX)
                Dim angleSensAPBc = Math.Atan2(vsAPBc.ProperY, vsAPBc.ProperX)

                Dim angleSensA = RoundToSignificantDigits(angleSensAPAb - angleSensAPBc, RealPrecision)

                While angleSensA < 0
                    angleSensA += 2 * Math.PI
                End While

                angleSensA *= RoundToSignificantDigits(180 / Math.PI, RealPrecision)
                sensAAngles.Add(angleSensA)

                Dim angleSensB = 360.0 - angleSensA
                While angleSensB < 0
                    angleSensB += 360
                End While
                sensBAngles.Add(angleSensB)
            Next

            For i = 0 To sensAAngles.Count - 1
                sommeAnglesSensA += sensAAngles(i)
                sommeAnglesSensB += sensBAngles(i)
            Next

            If sommeAnglesSensA > sommeAnglesSensB Then
                SensDeLecture = False
                Return New Tuple(Of List(Of Double), Boolean)(sensBAngles, False)
            Else
                SensDeLecture = True
                Return New Tuple(Of List(Of Double), Boolean)(sensAAngles, True)
            End If
        Else
            Throw New NotImplementedException("Sorry only polygons have angles.")
        End If
    End Function

    'Etrange Fonction je ne comprend pas comment elle fonctionne et son objectif
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Function GetProjectedInnerWideAngles() As Dictionary(Of Integer, Tuple(Of Point, Point))
        Dim result As New Dictionary(Of Integer, Tuple(Of Point, Point))
        Dim innerAngles = GetInnerAngles()
        Dim angles = innerAngles.Item1
        Dim sensPolygon = innerAngles.Item2
        Dim indexesOfPointsToProject As New List(Of Integer)
        Dim projectedPoints As New List(Of Tuple(Of Point, Point))
        For i = 0 To angles.Count - 1
            Dim angle = angles(i)
            If angle > 180 Then
                indexesOfPointsToProject.Add(i)
            End If
        Next
        For Each indexOfPointToProject In indexesOfPointsToProject
            Dim pA, pB, pC As Point
            If indexOfPointToProject = 0 Then
                pA = Points.Last
                pB = Points(indexOfPointToProject)
                pC = Points(indexOfPointToProject + 1)
            ElseIf indexOfPointToProject = Points.Count - 1 Then
                pA = Points(indexOfPointToProject - 1)
                pB = Points(indexOfPointToProject)
                pC = Points.First
            Else
                pA = Points(indexOfPointToProject - 1)
                pB = Points(indexOfPointToProject)
                pC = Points(indexOfPointToProject + 1)
            End If

            Dim projectedAb As Point = Nothing, projectedCb As Point = Nothing

            For i = 0 To Points.Count - 1
                If Not (i <= indexOfPointToProject + 1 And i >= indexOfPointToProject - 1) Then
                    Dim p1, p2 As Point
                    If i = 0 Then
                        p1 = Points.Last
                        p2 = Points.First
                    Else
                        p1 = Points(i - 1)
                        p2 = Points(i)
                    End If
                    Dim errorCode As ErrorCode = ErrorCode.NoError

                    Dim potentialProjectedAb As Point = GetIntersection(pA, pB, p1, p2, errorCode, True)
                    If errorCode = ErrorCode.NoError Then
                        If Distance(potentialProjectedAb, p1, p2, False) = 0 And Distance(potentialProjectedAb, pB) < Distance(potentialProjectedAb, pA) Then
                            If projectedAb Is Nothing Then
                                projectedAb = potentialProjectedAb
                            ElseIf Distance(projectedAb, pB) > Distance(potentialProjectedAb, pB) Then
                                projectedAb = potentialProjectedAb
                            End If
                        End If
                    End If

                    Dim potentialProjectedCb As Point = GetIntersection(pC, pB, p1, p2, errorCode, True)
                    If errorCode = ErrorCode.NoError Then
                        If Distance(potentialProjectedCb, p1, p2, False) = 0 And Distance(potentialProjectedCb, pB) < Distance(potentialProjectedCb, pC) Then
                            If projectedCb Is Nothing Then
                                projectedCb = potentialProjectedCb
                            ElseIf Distance(projectedCb, pB) > Distance(potentialProjectedCb, pB) Then
                                projectedCb = potentialProjectedCb
                            End If
                        End If
                    End If

                End If
            Next

            projectedPoints.Add(New Tuple(Of Point, Point)(projectedAb, projectedCb))
            result.Add(indexOfPointToProject, projectedPoints.Last)
        Next
        Return result
    End Function


    Public Function GetExtendedLines() As List(Of Tuple(Of Point, Point, Boolean))
        Dim angles = GetInnerAngles()
        Dim projectedInnerWideAngles = GetProjectedInnerWideAngles()
        Dim originalPolygon As New List(Of Tuple(Of Point, Point, Boolean))

        Dim sections As New List(Of Integer)
        Dim culsDeSac As New List(Of Integer)

        For i = 0 To Points.Count - 1
            Dim angle1, angle2 As Double
            If i = 0 Then
                angle1 = angles.Item1.Last
                angle2 = angles.Item1.First
            Else
                angle1 = angles.Item1(i - 1)
                angle2 = angles.Item1(i)
            End If

            If angle2 > 180 Then
                sections.Add(i)
            End If

            If Math.Round(angle1 + angle2, 0) = 180 Then
                culsDeSac.Add(i)
            End If
        Next

        Dim lastExtremityIndexes As New List(Of Integer)
        Dim lastDistance = Double.PositiveInfinity
        Dim enteredSection As Boolean = False

        If sections.Count = 0 Then
            enteredSection = True
        End If

        For i = 0 To Points.Count - 1
            Dim p1, p2 As Point

            If i = 0 Then
                p1 = Points.Last
                p2 = Points.First
            Else
                p1 = Points(i - 1)
                p2 = Points(i)
            End If

            If culsDeSac.Contains(i) And enteredSection Then
                Dim dP1P2 = Distance(p1, p2)
                If dP1P2 < lastDistance Then
                    lastDistance = dP1P2
                    If lastExtremityIndexes.Count > 0 Then
                        For Each lastExtremityIndex In lastExtremityIndexes
                            Dim lastExtremity = originalPolygon(lastExtremityIndex)
                            originalPolygon(lastExtremityIndex) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, False)
                        Next
                    End If
                    lastExtremityIndexes.Clear()
                    lastExtremityIndexes.Add(i)
                ElseIf dP1P2 = lastDistance Then
                    If lastExtremityIndexes.Contains(i - 1) Then
                        Dim lastExtremity = originalPolygon(i - 1)
                        originalPolygon(i - 1) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, False)
                        lastExtremityIndexes.Remove(i - 1)
                    End If
                    lastExtremityIndexes.Add(i)
                End If
            End If
            originalPolygon.Add(New Tuple(Of Point, Point, Boolean)(p1, p2, lastExtremityIndexes.Contains(i)))
            If sections.Contains(i) Then
                enteredSection = True
                lastDistance = Double.PositiveInfinity
                lastExtremityIndexes.Clear()
            End If
        Next

        If sections.Count > 0 Then
            For i = 0 To sections.First
                Dim p1, p2 As Point

                If i = 0 Then
                    p1 = Points.Last
                    p2 = Points.First
                Else
                    p1 = Points(i - 1)
                    p2 = Points(i)
                End If

                If culsDeSac.Contains(i) And enteredSection Then
                    Dim dP1P2 = Distance(p1, p2)
                    If dP1P2 < lastDistance Then
                        lastDistance = dP1P2
                        If lastExtremityIndexes.Count > 0 Then
                            For Each lastExtremityIndex In lastExtremityIndexes
                                Dim lastExtremity = originalPolygon(lastExtremityIndex)
                                originalPolygon(lastExtremityIndex) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, False)
                            Next
                        End If
                        lastExtremityIndexes.Clear()
                        lastExtremityIndexes.Add(i)
                    ElseIf dP1P2 = lastDistance Then
                        If lastExtremityIndexes.Contains(i - 1) Then
                            Dim lastExtremity = originalPolygon(i - 1)
                            originalPolygon(i - 1) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, False)
                            lastExtremityIndexes.Remove(i - 1)
                        End If
                        lastExtremityIndexes.Add(i)
                    End If
                End If
                originalPolygon(i) = New Tuple(Of Point, Point, Boolean)(p1, p2, lastExtremityIndexes.Contains(i))
            Next
        End If

        If lastExtremityIndexes.Count = 1 Then
            Dim beauGosse = lastExtremityIndexes.First
            lastExtremityIndexes.Clear()
            lastExtremityIndexes.Add(beauGosse)
            For i = 0 To Points.Count - 1
                If i <> beauGosse Then
                    Dim p1, p2 As Point

                    If i = 0 Then
                        p1 = Points.Last
                        p2 = Points.First
                    Else
                        p1 = Points(i - 1)
                        p2 = Points(i)
                    End If

                    If culsDeSac.Contains(i) And enteredSection Then
                        Dim dP1P2 = Distance(p1, p2)
                        If dP1P2 < lastDistance Then
                            lastDistance = dP1P2
                            If lastExtremityIndexes.Count > 0 Then
                                For Each lastExtremityIndex In lastExtremityIndexes
                                    Dim lastExtremity = originalPolygon(lastExtremityIndex)
                                    originalPolygon(lastExtremityIndex) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, i = beauGosse)
                                Next
                            End If
                            lastExtremityIndexes.Clear()
                            lastExtremityIndexes.Add(beauGosse)
                            lastExtremityIndexes.Add(i)
                        ElseIf dP1P2 = lastDistance Then
                            If lastExtremityIndexes.Contains(i - 1) Then
                                Dim lastExtremity = originalPolygon(i - 1)
                                originalPolygon(i - 1) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, i = beauGosse)
                                lastExtremityIndexes.Remove(i - 1)
                            End If
                            lastExtremityIndexes.Add(beauGosse)
                            lastExtremityIndexes.Add(i)
                        End If
                    End If
                    If sections.Contains(i) Then
                        enteredSection = True
                        lastDistance = Double.PositiveInfinity
                        lastExtremityIndexes.Clear()
                        lastExtremityIndexes.Add(beauGosse)
                    End If
                End If
            Next
            If sections.Count > 0 Then
                For i = 0 To sections.First
                    If i <> beauGosse Then
                        Dim p1, p2 As Point

                        If i = 0 Then
                            p1 = Points.Last
                            p2 = Points.First
                        Else
                            p1 = Points(i - 1)
                            p2 = Points(i)
                        End If

                        If culsDeSac.Contains(i) And enteredSection Then
                            Dim dP1P2 = Distance(p1, p2)
                            If dP1P2 < lastDistance Then
                                lastDistance = dP1P2
                                If lastExtremityIndexes.Count > 0 Then
                                    For Each lastExtremityIndex In lastExtremityIndexes
                                        Dim lastExtremity = originalPolygon(lastExtremityIndex)
                                        originalPolygon(lastExtremityIndex) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, i = beauGosse)
                                    Next
                                End If
                                lastExtremityIndexes.Clear()
                                lastExtremityIndexes.Add(beauGosse)
                                lastExtremityIndexes.Add(i)
                            ElseIf dP1P2 = lastDistance Then
                                If lastExtremityIndexes.Contains(i - 1) Then
                                    Dim lastExtremity = originalPolygon(i - 1)
                                    originalPolygon(i - 1) = New Tuple(Of Point, Point, Boolean)(lastExtremity.Item1, lastExtremity.Item2, i = beauGosse)
                                    lastExtremityIndexes.Remove(i - 1)
                                End If
                                lastExtremityIndexes.Add(beauGosse)
                                lastExtremityIndexes.Add(i)
                            End If
                        End If
                        originalPolygon(i) = New Tuple(Of Point, Point, Boolean)(p1, p2, lastExtremityIndexes.Contains(i))
                    End If
                Next
            End If
        End If

        For Each projectedInnerWideAngle In projectedInnerWideAngles
            Dim segmentAb = originalPolygon(projectedInnerWideAngle.Key)
            Dim projectedAb = projectedInnerWideAngle.Value.Item1

            If Distance(segmentAb.Item1, projectedAb) > Distance(segmentAb.Item2, projectedAb) Then
                originalPolygon(projectedInnerWideAngle.Key) = New Tuple(Of Point, Point, Boolean)(segmentAb.Item1, projectedAb, segmentAb.Item3)
            Else
                originalPolygon(projectedInnerWideAngle.Key) = New Tuple(Of Point, Point, Boolean)(projectedAb, segmentAb.Item2, segmentAb.Item3)
            End If
            Dim bcIndex = projectedInnerWideAngle.Key + 1
            If bcIndex >= originalPolygon.Count Then
                bcIndex -= originalPolygon.Count
            End If
            Dim segmentBc = originalPolygon(bcIndex)
            Dim projectedBc = projectedInnerWideAngle.Value.Item2

            If Distance(segmentBc.Item1, projectedBc) > Distance(segmentBc.Item2, projectedBc) Then
                originalPolygon(bcIndex) = New Tuple(Of Point, Point, Boolean)(segmentBc.Item1, projectedBc, segmentBc.Item3)
            Else
                originalPolygon(bcIndex) = New Tuple(Of Point, Point, Boolean)(projectedBc, segmentBc.Item2, segmentBc.Item3)
            End If
        Next

        Return originalPolygon
    End Function

    Public Function GetMergedExtendedLines() As List(Of Tuple(Of Point, Point, Vector, Boolean))
        Dim extendedLines = GetExtendedLines()
        Dim fusedLines As New List(Of Tuple(Of Point, Point, Vector, Boolean))
        Dim indexOfFusedLines As New List(Of Integer)
        Dim normales = GetNormales()
        For i = 0 To extendedLines.Count - 1
            If Not indexOfFusedLines.Contains(i) Then
                Dim normale = normales(i)
                Dim extendedLine = extendedLines(i)
                Dim currentFusedLine As Tuple(Of Point, Point, Boolean) = Nothing
                For j = 0 To extendedLines.Count - 1
                    If i <> j And Not indexOfFusedLines.Contains(j) Then
                        Dim potentialFusableLine = extendedLines(j)
                        If New Vector(extendedLine.Item1, extendedLine.Item2).ScaleOneVector = New Vector(potentialFusableLine.Item1, potentialFusableLine.Item2).ScaleOneVector Or New Vector(extendedLine.Item1, extendedLine.Item2).ScaleOneVector = New Vector(potentialFusableLine.Item1, potentialFusableLine.Item2).ScaleOneVector.Reverse Then
                            If Distance(extendedLine.Item1, potentialFusableLine.Item1, potentialFusableLine.Item2, False) = 0.0 Or
                                    Distance(extendedLine.Item2, potentialFusableLine.Item1, potentialFusableLine.Item2, False) = 0.0 Or
                                    Distance(potentialFusableLine.Item1, extendedLine.Item1, extendedLine.Item2, False) = 0.0 Or
                                    Distance(potentialFusableLine.Item2, extendedLine.Item1, extendedLine.Item2, False) = 0.0 Then

                                Dim dico As New Dictionary(Of Tuple(Of Point, Point, Boolean), Double)

                                Dim ab = Distance(extendedLine.Item1, extendedLine.Item2), cd = Distance(potentialFusableLine.Item1, potentialFusableLine.Item2),
                                    ac = Distance(extendedLine.Item1, potentialFusableLine.Item1), bd = Distance(extendedLine.Item2, potentialFusableLine.Item2),
                                    bc = Distance(extendedLine.Item2, potentialFusableLine.Item1), ad = Distance(extendedLine.Item1, potentialFusableLine.Item2)

                                dico.Add(extendedLine, ab)
                                dico.Add(potentialFusableLine, cd)
                                dico.Add(New Tuple(Of Point, Point, Boolean)(extendedLine.Item1, potentialFusableLine.Item1, extendedLine.Item3), ac)
                                dico.Add(New Tuple(Of Point, Point, Boolean)(extendedLine.Item2, potentialFusableLine.Item2, extendedLine.Item3), bd)
                                dico.Add(New Tuple(Of Point, Point, Boolean)(extendedLine.Item2, potentialFusableLine.Item1, extendedLine.Item3), bc)
                                Try
                                    dico.Add(New Tuple(Of Point, Point, Boolean)(extendedLine.Item1, potentialFusableLine.Item2, extendedLine.Item3), ad)
                                Catch
                                End Try

                                Dim max As Double = -1
                                Dim maxIndex As Integer = -1

                                For k = 0 To dico.Values.Count - 1
                                    Dim value = dico.Values(k)
                                    If value > max Then
                                        max = value
                                        maxIndex = k
                                    End If
                                Next

                                indexOfFusedLines.Add(j)
                                indexOfFusedLines.Add(i)

                                currentFusedLine = dico.Keys(maxIndex)
                                extendedLine = currentFusedLine

                                j = -1
                            End If
                        End If
                    End If
                Next
                If currentFusedLine Is Nothing Then
                    fusedLines.Add(New Tuple(Of Point, Point, Vector, Boolean)(extendedLine.Item1, extendedLine.Item2, normale, extendedLine.Item3))
                Else
                    fusedLines.Add(New Tuple(Of Point, Point, Vector, Boolean)(currentFusedLine.Item1, currentFusedLine.Item2, normale, extendedLine.Item3))
                End If
            End If
        Next
        Return fusedLines
    End Function

    ''' <summary>
    ''' Get the normal vectors of polygon edges
    ''' </summary>
    ''' <returns> List of vectors </returns>
    Public Function GetNormales() As List(Of Vector)
        If Type = SurfaceType.Rectangle OrElse Type = SurfaceType.Polygon Then
            Dim normales As New List(Of Vector)
            If SensDeLecture Is Nothing Then
                GetInnerAngles()
            End If

            For i = 0 To Points.Count - 1
                Dim p1, p2 As Point
                If i = 0 Then
                    p1 = Points.Last
                    p2 = Points.First
                Else
                    p1 = Points(i - 1)
                    p2 = Points(i)
                End If
                Dim v = New Vector(New Point(p1.X, p1.Y), New Point(p2.X, p2.Y)).ScaleOneVector
                Dim n = New Vector() With {.X = -v.Y, .Y = v.X}
                n.Origin = MidPoint(p1, p2)
                If SensDeLecture Then
                    normales.Add(n)
                Else
                    normales.Add(n.Reverse)
                End If
            Next
            Return normales
        Else
            Throw New NotImplementedException("Not Implemented for this type of Surface")
        End If
    End Function

    ''' <summary>
    ''' Get Normal the face. Use CrossProduct between 2 edges vector.
    ''' </summary>
    ''' <returns>Trigonometric normale</returns>
    Public Function GetNormal() As Vector
        If Type = SurfaceType.Polygon OrElse Type = SurfaceType.Rectangle Then
            'Setting approximation
            'Dim previousApproximateCalculus = Vector.ApproximateCalculus
            'Dim previousPrecision = Vector.Precision
            'Vector.ApproximateCalculus = True
            'Vector.Precision = 3
            'Computing 2 edges vectors
            Dim firstLine As Line = Edges.FirstOrDefault()
            Dim firstVector As Vector = firstLine.Vector.ScaleOneVector
            Dim secondLine As Line = Edges(1)
            Dim secondVector As Vector = secondLine.Vector.ScaleOneVector

            'find 2 non null vector
            Dim i = 1
            While (firstVector.X = 0 AndAlso firstVector.Y = 0 AndAlso firstVector.Z = 0 OrElse secondVector.X = 0 AndAlso secondVector.Y = 0 AndAlso secondVector.Z = 0) AndAlso i + 1 < Edges.Count
                firstLine = Edges(i)
                firstVector = firstLine.Vector.ScaleOneVector
                secondLine = Edges(i + 1)
                secondVector = secondLine.Vector.ScaleOneVector
                i += 1
                If i + 1 > Edges.Count - 1 Then
                    Exit While
                End If
            End While


            If firstVector.X = 0 AndAlso firstVector.Y = 0 AndAlso firstVector.Z = 0 OrElse secondVector.X = 0 AndAlso secondVector.Y = 0 AndAlso secondVector.Z = 0 Then
                Return Nothing
            End If

            Dim counter = 1
            'Creating additional condition to find 2 correct edges
            Dim additionalCondition = Nothing
            'If we are in trigo sens, The angle between the 2 edges must be positive.
            If IsSensTrigo() Then
                additionalCondition = secondVector.Angle(firstVector, True) < 0
            Else
                additionalCondition = secondVector.Angle(firstVector, True) > 0
            End If
            'less than 2 edges?
            If counter >= Edges.Count Then
                'This is a line. exit
                'reset approx
                'Vector.ApproximateCalculus = previousApproximateCalculus
                'Vector.Precision = previousPrecision
                Return Nothing
            End If
            'Search for 2 correct vectors. Must not be aligned. Must not open the polygon. For exemple, we must not take the "internal" edges of an "L" shape polygon
            While additionalCondition OrElse (secondVector.Angle(firstVector, True) > -1 And secondVector.Angle(firstVector, True) < 1 Or secondVector.Angle(firstVector, True) > 179 And secondVector.Angle(firstVector, True) < 181)
                'same vector or direction, or opening angle. we need to find others.
                If counter > Edges.Count Then
                    'counter exceed edge count
                    'No correct angle found.
                    'Vector.ApproximateCalculus = previousApproximateCalculus
                    'Vector.Precision = previousPrecision
                    Return Nothing
                End If

                'update vectors
                firstVector = New Vector(secondVector.X, secondVector.Y, secondVector.Z)
                If counter = Edges.Count Then
                    secondLine = Edges(0)
                Else
                    secondLine = Edges(counter)
                End If
                secondVector = secondLine.Vector.ScaleOneVector

                'Update condition
                If IsSensTrigo() Then
                    additionalCondition = secondVector.Angle(firstVector, True) < 0
                Else
                    additionalCondition = secondVector.Angle(firstVector, True) > 0
                End If
                counter += 1
            End While
            'reset approx
            'Vector.ApproximateCalculus = previousApproximateCalculus
            'Vector.Precision = previousPrecision
            Return Vector.CrossProduct(firstVector, secondVector).ScaleOneVector
        Else
            Throw New NotImplementedException("Only for polygon and rectangle")
        End If
    End Function

    ''' <summary>
    ''' Compute angle of edge 2 by 2. Sum them. if 360° then True, else -360 then false.
    ''' </summary>
    ''' <returns></returns>
    Public Function IsSensTrigo() As Boolean
        Dim angleSum = 0.0
        Dim positivCounter = 0
        Dim negativCounter = 0
        For i As Integer = 1 To Edges.Count
            Dim firstLine As Line = Edges(i - 1)
            Dim firstVector As Vector = firstLine.Vector.ScaleOneVector
            Dim secondLine As Line = Nothing
            If i = Edges.Count Then
                secondLine = Edges(0)
            Else
                secondLine = Edges(i)
            End If
            Dim secondVector As Vector = secondLine.Vector.ScaleOneVector
            If firstVector.X = 0 AndAlso firstVector.Y = 0 AndAlso firstVector.Z = 0 OrElse secondVector.X = 0 AndAlso secondVector.Y = 0 AndAlso secondVector.Z = 0 Then
                Continue For
            End If
            Dim curAngle = secondVector.Angle(firstVector, True)
            angleSum += curAngle
            If curAngle >= 0 Then
                positivCounter += 1
            Else
                negativCounter += 1
            End If
        Next
        If angleSum > 359 Then
            Return True
        ElseIf angleSum < -359 Then
            Return False
        Else
            'There is a problem. Maybe a 180° angle.
            If positivCounter >= negativCounter Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    '''' <summary>
    '''' Constructs Axis (for walls) from Surface representing footprint of walls
    '''' </summary>
    '''' First Version, prefer V2 and V3
    '''' <returns></returns>
    'Public Function GetAxis() As IReadOnlyList(Of Tuple(Of Tuple(Of Point, Point, Vector, Boolean), Tuple(Of Point, Point, Vector, Boolean)))
    '    Dim lines = GetMergedExtendedLines()
    '    Dim couples As New List(Of Tuple(Of Tuple(Of Point, Point, Vector, Boolean), Tuple(Of Point, Point, Vector, Boolean)))
    '    Dim dimensions As New Dictionary(Of Vector, List(Of Tuple(Of Point, Point, Vector, Boolean)))
    '    For Each line In lines
    '        If Not line.Item4 Then
    '            Dim normale = line.Item3
    '            Dim reversedNormale = normale.Reverse
    '            Dim indexOf = -1
    '            For i = 0 To dimensions.Count - 1
    '                Dim currentKey = dimensions.Keys(i)
    '                If currentKey = normale Or currentKey = reversedNormale Then
    '                    indexOf = i
    '                    Exit For
    '                End If
    '            Next
    '            If indexOf = -1 Then
    '                dimensions.Add(normale, New List(Of Tuple(Of Point, Point, Vector, Boolean)))
    '                indexOf = dimensions.Count - 1
    '            End If
    '            dimensions.ElementAt(indexOf).Value.Add(line)
    '        End If
    '    Next
    '    For Each dimension In dimensions
    '        Dim gentlemen As New Dictionary(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)),
    '            ladies As New List(Of Tuple(Of Point, Point, Vector, Boolean))
    '        Dim reference = dimension.Key
    '        For Each tuple In dimension.Value
    '            If tuple.Item3 = reference Then
    '                gentlemen.Add(tuple, New List(Of WillToDance))
    '            Else
    '                ladies.Add(tuple)
    '            End If
    '        Next
    '        For Each gentleman In gentlemen
    '            For Each lady In ladies
    '                Dim projectedLadyOrigin = GetNormalPoint(lady.Item1, gentleman.Key.Item3.Origin, gentleman.Key.Item3.EndPoint)
    '                Dim toCheck = New Vector(gentleman.Key.Item3.Origin, projectedLadyOrigin).ScaleOneVector
    '                If gentleman.Key.Item3.ScaleOneVector = toCheck Then
    '                    Dim newWillToDance As New WillToDance() With {.Lady = lady, .Will = Distance(New Vector(gentleman.Key.Item1, gentleman.Key.Item2),
    '                                                                                          New Vector(lady.Item1, lady.Item2)), .TakenBy = Nothing}
    '                    gentleman.Value.Add(newWillToDance)
    '                End If
    '            Next
    '            gentleman.Value.Sort(Function(lady1 As WillToDance, lady2 As WillToDance) As Integer
    '                                     Return lady1.Will.CompareTo(lady2.Will)
    '                                 End Function)
    '        Next
    '        If ladies.Count > 0 Then
    '            While Not IsEveryoneSatisfied(gentlemen)
    '                Dim mostWillingGentleman = GetMostWillingGentleman(gentlemen)
    '                FightForYourLady(mostWillingGentleman, gentlemen, ladies)
    '            End While

    '            For Each gentleman In gentlemen
    '                Dim gentlemanHasALady As Boolean = False
    '                For Each will In gentleman.Value
    '                    If will.TakenBy IsNot Nothing Then ' A VERIFIER !!
    '                        If IsItTheSamePerson(will.TakenBy.Value.Key, gentleman.Key) Then
    '                            couples.Add(New Tuple(Of Tuple(Of Point, Point, Vector, Boolean), Tuple(Of Point, Point, Vector, Boolean))(gentleman.Key, will.Lady))
    '                            gentlemanHasALady = True
    '                        End If
    '                    End If
    '                Next
    '                If Not gentlemanHasALady And gentleman.Value.Count > 0 Then

    '                End If
    '            Next
    '        End If
    '    Next

    '    Return couples
    'End Function

#Region "GetAxis Functions"
    Private Sub FightForYourLady(gentleman As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)),
                                 gentlemen As Dictionary(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)),
                                 ladies As List(Of Tuple(Of Point, Point, Vector, Boolean)))
        Dim gotMyFirstChoice As Boolean? = Nothing
        Dim aLadyHasChosenThisGentleman As Boolean = False
        Dim letItGo As Boolean = False
        Dim rivals As New List(Of KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)))

        Dim ladiesOutOfLeague As New List(Of WillToDance)

        While (Not aLadyHasChosenThisGentleman And gentleman.Value.Count > 0) And (Not letItGo)
            For Each will In gentleman.Value
                If will.TakenBy Is Nothing Then
                    If gotMyFirstChoice Is Nothing Then
                        gotMyFirstChoice = True
                        TellEveryoneThisLadyHasAcceptedMe(will.Lady, gentleman, gentlemen)
                        aLadyHasChosenThisGentleman = True
                        Exit For
                    ElseIf gotMyFirstChoice = False Then
                        Dim aRivalWantedHerMoreThanHim As Boolean = False
                        For Each rival In rivals
                            If HowMuchDoesHeWillsToDanceWithHer(rival, will.Lady) < will.Will Then
                                aRivalWantedHerMoreThanHim = True
                                For Each rivalWill In rival.Value
                                    If rivalWill.TakenBy IsNot Nothing Then
                                        If IsItTheSamePerson(rivalWill.TakenBy.Value.Key, rival.Key) Then
                                            TellEveryoneThisLadyHasAcceptedMe(rivalWill.Lady, Nothing, gentlemen)
                                            TellEveryoneThisLadyHasAcceptedMe(will.Lady, rival, gentlemen)
                                            Exit For
                                        End If
                                    End If
                                Next
                                Exit For
                            End If
                        Next
                        If aRivalWantedHerMoreThanHim Then
                            Exit For
                        Else
                            TellEveryoneThisLadyHasAcceptedMe(will.Lady, gentleman, gentlemen)
                            aLadyHasChosenThisGentleman = True
                            Exit For
                        End If
                    ElseIf gotMyFirstChoice = True Then
                        Throw New InvalidOperationException("Why are you searching for a lady if you already got one?!")
                    End If
                Else
                    If HowMuchDoesHeWillsToDanceWithHer(will.TakenBy.Value, will.Lady) < will.Will Then
                        ladiesOutOfLeague.Add(will)
                        will.Rejected = True
                    Else
                        rivals.Add(will.TakenBy)
                    End If
                End If
                gotMyFirstChoice = False
                If ladiesOutOfLeague.Count = ladies.Count Then
                    letItGo = True
                End If
            Next
            rivals.Clear()
        End While
        If letItGo Then
            For i = 0 To gentleman.Value.Count - 1
                Dim will = gentleman.Value(i)
                will.Rejected = True
                gentleman.Value(i) = will
            Next
        End If
    End Sub

    Private Sub TellEveryoneThisLadyHasAcceptedMe(lady As Tuple(Of Point, Point, Vector, Boolean), happyOne As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))?, gentlemen As Dictionary(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)))
        For Each gentleman In gentlemen
            Dim index = 0
            For Each willingLady In gentleman.Value
                If IsItTheSamePerson(willingLady.Lady, lady) Then
                    gentleman.Value(index) = New WillToDance() With {.Lady = lady, .TakenBy = happyOne, .Will = willingLady.Will}
                    willingLady.TakenBy = happyOne
                    Exit For
                End If
                index += 1
            Next
        Next
    End Sub

    Private Function IsItTheSamePerson(person1 As Tuple(Of Point, Point, Vector, Boolean), person2 As Tuple(Of Point, Point, Vector, Boolean)) As Boolean
        Return person1.Item1 = person2.Item1 And person1.Item2 = person2.Item2 And person1.Item3 = person2.Item3 And person1.Item4 = person2.Item4
    End Function

    Private Function GetMostWillingGentleman(gentlemen As Dictionary(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))) As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))
        Dim mostWillingDude As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))? = Nothing
        For Each gentleman In gentlemen
            If Not HasLadyAcceptedThisGentleman(gentleman) Then
                If mostWillingDude Is Nothing Then
                    mostWillingDude = gentleman
                Else
                    mostWillingDude = WhoWillsTheMost(mostWillingDude, gentleman)
                End If
            End If
        Next
        Return mostWillingDude
    End Function

    Private Function HowMuchDoesHeWillsToDanceWithHer(gentleman As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)), lady As Tuple(Of Point, Point, Vector, Boolean)) As Double
        For Each will In gentleman.Value
            If IsItTheSamePerson(will.Lady, lady) Then
                Return will.Will
            End If
        Next
        Return Double.PositiveInfinity
    End Function

    Private Function WhoWillsTheMost(suitor1 As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance)), suitor2 As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))) As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))
        Dim suitor1BestWill As Double = Double.PositiveInfinity, suitor2BestWill As Double = Double.PositiveInfinity
        For Each lady In suitor1.Value
            If suitor1BestWill > lady.Will Then
                suitor1BestWill = lady.Will
            End If
        Next

        For Each lady In suitor2.Value
            If suitor2BestWill > lady.Will Then
                suitor2BestWill = lady.Will
            End If
        Next

        If suitor1BestWill < suitor2BestWill Then
            Return suitor1
        Else
            Return suitor2
        End If
    End Function

    Private Function IsEveryoneSatisfied(gentlemen As Dictionary(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))) As Boolean
        For Each gentleman In gentlemen
            If gentleman.Value.Count > 0 Then
                If Not HasLadyAcceptedThisGentleman(gentleman) Then
                    Return False
                End If
            End If
        Next
        Return True
    End Function

    Private Function HasLadyAcceptedThisGentleman(gentleman As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))) As Boolean
        Dim satisfied As Boolean = False
        Dim hasGiveUp As Boolean = True
        For Each lady In gentleman.Value
            If lady.Rejected = False Then
                hasGiveUp = False
                Exit For
            End If
        Next
        If hasGiveUp Then
            satisfied = True
        Else
            For Each lady In gentleman.Value
                If lady.TakenBy IsNot Nothing Then
                    If IsItTheSamePerson(lady.TakenBy.Value.Key, gentleman.Key) Then
                        satisfied = True
                        Exit For
                    End If
                End If
            Next
        End If
        Return satisfied
    End Function

    Private Structure WillToDance
        Public Lady As Tuple(Of Point, Point, Vector, Boolean)
        Public Will As Double
        Public TakenBy As KeyValuePair(Of Tuple(Of Point, Point, Vector, Boolean), List(Of WillToDance))?
        Public Rejected As Boolean
    End Structure
#End Region

    ''' <summary>
    ''' Constructs Axis (for walls) from Surface representing footprint of walls
    ''' </summary>
    ''' <param name="epMin"> Minimum thickness of walls </param>
    ''' <param name="epMax"> Maximum thockness of walls </param>
    ''' <param name="epList"> List of possible thickness for walls </param>
    ''' Second Version, a Third Version exist in WallAxis code
    ''' <returns></returns>
    Public Function GetAxisV2(epMin As Double, epMax As Double, Optional epList As List(Of Double) = Nothing) As List(Of Tuple(Of Line, Line))
        Dim angles = GetInnerAngles()
        Dim normales = GetNormales()
        Dim linesToKeep As New List(Of Line)
        Dim normalesToKeep As New List(Of Vector)
        Dim lgMax = Math.Round(epMax, 2)
        Dim lgMin = Math.Round(epMin, 2)
        Dim lgList = epList

        'On cherche les lignes du polygone pouvant etre des longueur de murs et non des épaisseurs
        'Pour cela on utilise la propriété qu'une ligne d'épaissseur ne pas pas avoir n'importe quelle longueur 
        'et que ses deux lignes voisines sont parallèles (la somme des angles aux etrémités de la ligne fait 180°
        For i = 0 To Points.Count - 1
            Dim angle1, angle2 As Double
            Dim p1, p2 As Point
            If i = 0 Then
                angle1 = angles.Item1.Last
                angle2 = angles.Item1(i)
                p1 = Points.Last
                p2 = Points(i)
            Else
                angle1 = angles.Item1(i - 1)
                angle2 = angles.Item1(i)
                p1 = Points(i - 1)
                p2 = Points(i)
            End If

            Dim dP1P2 = Distance(p1, p2)

            'On distingue le cas où une liste d'épaisseur précise a été définie par rapport à des simples min et max
            If lgList IsNot Nothing Then
                If Not (Math.Round(angle1 + angle2, 0) = 180) Or Not lgList.Contains(Math.Round(dP1P2, 2)) Then
                    linesToKeep.Add(New Line(p1, p2))
                    normalesToKeep.Add(normales(i))
                End If

            Else
                If Not (Math.Round(angle1 + angle2, 0) = 180) Or dP1P2 > lgMax Or dP1P2 < lgMin Then
                    linesToKeep.Add(New Line(p1, p2))
                    normalesToKeep.Add(normales(i))
                End If
            End If
        Next
        Dim coupleLines As New List(Of Tuple(Of Line, Line))

        'Création de la matrice d'associatiuon des lignes
        Dim isAssociated(linesToKeep.Count - 1, linesToKeep.Count - 1) As Boolean
        For i = 0 To linesToKeep.Count - 1
            For j = 0 To linesToKeep.Count - 1
                isAssociated(i, j) = False
            Next
        Next

        'On cherche à associer à chaque ligne une autre ligne pour former les nus extérieurs d'un mur
        For j = 0 To linesToKeep.Count - 1
            Dim myLine = linesToKeep(j)
            Dim associatedLines As New List(Of Line)
            Dim lastDistance = Double.PositiveInfinity

            Dim indexesAssociated As New List(Of Integer)
            For k = 0 To linesToKeep.Count - 1
                If k <> j And Not isAssociated(k, j) Then
                    Dim testLine = linesToKeep(k)
                    Dim angle = myLine.Vector.Angle(testLine.Vector, True)

                    'Test de parallélisme
                    If (angle < 0.1 And angle > -0.1) Or (angle < 180.1 And angle > 179.9) Or (angle < 360.1 And angle > 359.9) Then
                        'Dim d_l1_l2 = Distance(line, testLine, False)
                        Dim commonProjectedLines = GetCommonProjectedLines(myLine, testLine)

                        'Test de recouvrement
                        If commonProjectedLines.Item1 IsNot Nothing And commonProjectedLines.Item2 IsNot Nothing Then
                            Dim dL1L2 = Distance(myLine, testLine, False)
                            Dim dir = New Vector(myLine.StartPoint, testLine.StartPoint)

                            'On distingue le cas où une liste d'épaisseur précise a été définie par rapport à des simples min et max
                            If lgList IsNot Nothing AndAlso lgList.Count > 0 Then
                                If lgList.Contains(Math.Round(dL1L2, 2)) And normalesToKeep(j) * dir > 0 Then
                                    If dL1L2 < lastDistance Then
                                        lastDistance = dL1L2

                                        associatedLines.Clear()
                                        associatedLines.Add(testLine)

                                        indexesAssociated.Clear()
                                        indexesAssociated.Add(k)

                                    ElseIf dL1L2 = lastDistance Then
                                        associatedLines.Add(testLine)
                                        indexesAssociated.Add(k)
                                    End If
                                End If
                            ElseIf dL1L2 <= lgMax And dL1L2 >= lgMin And normalesToKeep(j) * dir > 0 Then
                                If dL1L2 < lastDistance Then
                                    lastDistance = dL1L2

                                    associatedLines.Clear()
                                    associatedLines.Add(testLine)

                                    indexesAssociated.Clear()
                                    indexesAssociated.Add(k)

                                ElseIf dL1L2 = lastDistance Then
                                    associatedLines.Add(testLine)
                                    indexesAssociated.Add(k)
                                End If

                            End If

                        End If

                    End If

                End If
            Next

            If indexesAssociated.Count <> 0 Then
                For Each index In indexesAssociated
                    isAssociated(j, index) = True
                Next
            End If
            If associatedLines.Count <> 0 Then
                For Each associatedLine In associatedLines
                    coupleLines.Add(GetCommonProjectedLines(myLine, associatedLine))
                Next
            End If
        Next
        Return coupleLines
    End Function

    ''' <summary>
    ''' Get the gravity center of the Surface
    ''' </summary>
    ''' <returns> Gravity point </returns>
    Public ReadOnly Property GravityCenter As Point
        Get
            Select Case Type
                Case SurfaceType.Circle
                    Return Points(0)
                Case SurfaceType.Ellipse
                    Return Points(0)
                Case SurfaceType.Rectangle
                    Return MidPoint(Points(0), Points(2))
                Case SurfaceType.Polygon
                    Dim gx As Double = (Points.First.X + Points.Last.X) * (Points.Last.X * Points.First.Y - Points.First.X * Points.Last.Y)
                    Dim gy As Double = (Points.First.Y + Points.Last.Y) * (Points.Last.X * Points.First.Y - Points.First.X * Points.Last.Y)
                    If SensDeLecture Is Nothing Then
                        GetInnerAngles()
                    End If
                    For j = 0 To Points.Count - 2
                        gx += (Points(j).X + Points(j + 1).X) * (Points(j).X * Points(j + 1).Y - Points(j + 1).X * Points(j).Y)
                        gy += (Points(j).Y + Points(j + 1).Y) * (Points(j).X * Points(j + 1).Y - Points(j + 1).X * Points(j).Y)
                    Next
                    If SensDeLecture Then
                        Return New Point(gx / (6 * InternArea()), gy / (6 * InternArea()))
                    Else
                        Return New Point(-gx / (6 * InternArea()), -gy / (6 * InternArea()))
                    End If
                Case Else
                    Throw New Exception("Define surface type before checking area")
            End Select
        End Get
    End Property

    ''' <summary>
    ''' Integer representing the total number of times that Surface perimeter travel around a point
    ''' </summary>
    ''' <param name="testPoint"> Tested Point </param>
    ''' <returns></returns>
    Public Function WindingNumber(testPoint As Point) As Integer
        Dim wn As Integer = 0
        If Type = SurfaceType.Circle Then
            If Distance(Points(0), testPoint) <= Distance(Points(0), Points(1)) Then
                wn = 1
            End If
        ElseIf Type = SurfaceType.Ellipse Then
            Dim rayonA = Distance(Points(0), Points(1))
            Dim rayonB = Distance(Points(0), Points(2))
            Dim fPt, fPrime As New Point
            If rayonA < rayonB Then
                Dim temp = rayonA
                rayonA = rayonB
                rayonB = temp
                fPt = New Vector(Points(0), Points(2)).GetPointAtDistance(Math.Sqrt(rayonA * rayonA - rayonB * rayonB), Points(0))
                fPrime = New Vector(Points(0), Points(2)).GetPointAtDistance(-Math.Sqrt(rayonA * rayonA - rayonB * rayonB), Points(0))
            Else
                fPt = New Vector(Points(0), Points(1)).GetPointAtDistance(Math.Sqrt(rayonA * rayonA - rayonB * rayonB), Points(0))
                fPrime = New Vector(Points(0), Points(1)).GetPointAtDistance(-Math.Sqrt(rayonA * rayonA - rayonB * rayonB), Points(0))
            End If
            If Distance(fPt, testPoint) + Distance(fPrime, testPoint) <= 2 * rayonA Then
                wn = 1
            End If
        Else
            Dim sum As Double = 0.0
            For i = 0 To Points.Count - 1
                Dim cP As Point = Points(i)
                Dim sP As Point
                If i = Points.Count - 1 Then
                    sP = Points.First
                Else
                    sP = Points(i + 1)
                End If

                If cP = sP Then
                    Continue For
                End If

                If Distance(testPoint, cP, sP, False) = 0 Then
                    wn = 1
                    GoTo pointDedans
                End If

                Dim vCP = New Vector(testPoint, cP)
                Dim vSP = New Vector(testPoint, sP)

                Dim angleCP = Math.Atan2(vCP.ProperY, vCP.ProperX)
                Dim angleSP = Math.Atan2(vSP.ProperY, vSP.ProperX)
                Dim angleCPSP = angleSP - angleCP

                If angleCPSP < -Math.PI Then
                    angleCPSP += 2 * Math.PI
                End If
                If angleCPSP > Math.PI Then
                    angleCPSP -= 2 * Math.PI
                End If

                sum += angleCPSP

            Next
            wn = Math.Abs(Math.Round(sum / (2 * Math.PI)))
        End If
pointDedans:
        Return wn
    End Function

    ''' <summary>
    ''' Check if a 3D-Point is in a 3D-Surface, works only with Polygon and Rectangle
    ''' </summary>
    ''' <param name="testPoint"></param>
    ''' <param name="tolerance">Set the Tolerance as 10^-tolerance. Default is -4. Corresponds to maximal distance to plane.</param>
    ''' <returns></returns>
    Public Function IsPointInsideIn3D(testPoint As Point, Optional tolerance As Integer? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = DefaultTolerance

        Dim vectOne, vectTwo As Vector
        vectOne = New Vector(Me.Points(0), Me.Points(1))
        vectTwo = New Vector(Me.Points(1), Me.Points(2))
        Dim nullVector = New Vector
        Dim k = 2

        'Test d'appartenance au plan de la surface
        While Vector.CrossProduct(vectOne, vectTwo) = nullVector AndAlso k < Me.Points.Count
            If k + 1 > Me.Points.Count - 1 Then
                vectTwo = New Vector(Me.Points(k), Me.Points(0))
            Else
                vectTwo = New Vector(Me.Points(k), Me.Points(k + 1))
            End If
            k += 1
        End While
        Dim normalVector = Vector.CrossProduct(vectOne, vectTwo)
        If normalVector = nullVector Then
            Return False
        End If
        Dim mePlane = New Plane(Me.Points(0), normalVector.ScaleOneVector)
        If Not mePlane.DoesContainPoint(testPoint, tolerance) Then
            Return False
        End If

        If normalVector.X = 0 AndAlso normalVector.Y = 0 Then
            Return Me.WindingNumber(testPoint)
        End If

        'Changement de Repère
        'dim vectOneArray As Object() = {VectOne.X, VectOne.Y, VectOne.Z}
        'Dim normalVectorArray As Object() = {normalVector.X, normalVector.Y, normalVector.Z}
        'Treegram.GeomKernel.BasicModeler.Application.NormerVecteur(vectOneArray, vectOneArray)
        'Treegram.GeomKernel.BasicModeler.Application.NormerVecteur(normalVectorArray, normalVectorArray)

        'VectOne = VectOne.ScaleOneVector
        'normalVector = normalVector.ScaleOneVector
        Dim secondAxis = Vector.CrossProduct(normalVector, vectOne)
        'If secondAxis.X = 0 And secondAxis.Y = 0 And secondAxis.Z = 0 Then
        '    secondAxis = Vector.CrossProduct(normalVector, VectOne)
        'End If

        'Application.NormerVecteur(VectOne., VectOne)
        'Application.NormerVecteur(secondAxis, secondAxis)
        'Application.NormerVecteur(normalVector, normalVector)

        'VectOne = VectOne.ScaleOneVector
        'secondAxis = secondAxis.ScaleOneVector
        'normalVector = normalVector.ScaleOneVector
        Dim changeListPoints As New List(Of Point)
        If secondAxis.X = 0 And secondAxis.Y = 0 And secondAxis.Z = 0 Then
            secondAxis = Vector.CrossProduct(normalVector, vectOne)
        End If
        Dim cTp = ChangeCoordPointInAnotherAxisSyst({testPoint.X, testPoint.Y, testPoint.Z},
                                                                                     {Me.Points(0).X, Me.Points(0).Y, Me.Points(0).Z},
                                                                                     {vectOne.X, vectOne.Y, vectOne.Z},
                                                                                     {secondAxis.X, secondAxis.Y, secondAxis.Z},
                                                                                     {normalVector.X, normalVector.Y, normalVector.Z})
        Dim newTestPoint = New Point(cTp(0), cTp(1), 0)

        For Each pt In Me.Points
            Dim cP = ChangeCoordPointInAnotherAxisSyst({pt.X, pt.Y, pt.Z}, {Me.Points(0).X, Me.Points(0).Y, Me.Points(0).Z},
                                                                                     {vectOne.X, vectOne.Y, vectOne.Z},
                                                                                     {secondAxis.X, secondAxis.Y, secondAxis.Z},
                                                                                     {normalVector.X, normalVector.Y, normalVector.Z})
            changeListPoints.Add(New Point(cP(0), cP(1), 0))
        Next
        Dim newSurf = New Surface(changeListPoints, Me.Type)

        'WindingNumber
        Dim windingNumber As Integer = newSurf.WindingNumber(newTestPoint)
        Return windingNumber Mod 2 = 1
    End Function




    ''' <summary>
    ''' Offset of Surface perimeter
    ''' </summary>
    ''' Prefer the function Offset2
    ''' <param name="dist"> Distance of Offset</param>
    ''' <returns> new Surface </returns>
    Public Function Offset(dist As Double) As Surface
        If Type = SurfaceType.Circle Then
            Dim dir = New Vector(Points(0), Points(1))
            If dist >= dir.Length Then
                Return Nothing
            Else
                Dim offsetpoints As New List(Of Point) From {
                Points(0),
                New Point(Points(1).X + dist * dir.ProperX, Points(1).Y + dist * dir.ProperY)
            }
                Return New Surface(offsetpoints, Type)
            End If
        ElseIf Type = SurfaceType.Ellipse Then
            If dist >= Math.Min(Distance(Points(0), Points(1)), Distance(Points(0), Points(2))) Then
                Return Nothing
            Else
                Dim offsetpoints As New List(Of Point) From {
                    Points(0),
                    New Vector(Points(0), Points(1)).GetPointAtDistance(dist, Points(1)),
                    New Vector(Points(0), Points(2)).GetPointAtDistance(dist, Points(2))
                }
                Return New Surface(offsetpoints, Type)
            End If
        Else
            Dim lines As New List(Of Line) From {
                New Line(Points.Last, Points.First)
            }
            For i = 0 To Points.Count - 2
                lines.Add(New Line(Points(i), Points(i + 1)))
            Next

            Dim normales = GetNormales()

            Dim offsetLines As New List(Of Line)
            For j = 0 To lines.Count - 1
                Dim startP As Point = New Point(lines(j).StartPoint.X - dist * normales(j).X, lines(j).StartPoint.Y - dist * normales(j).Y)
                Dim endP As Point = New Point(lines(j).EndPoint.X - dist * normales(j).X, lines(j).EndPoint.Y - dist * normales(j).Y)
                offsetLines.Add(New Line(startP, endP))
            Next

            Dim offsetPoints As New List(Of Point)
            Dim erreur As ErrorCode

            For k = 0 To offsetLines.Count - 2
                Dim point = GetIntersection(offsetLines(k).StartPoint, offsetLines(k).EndPoint, offsetLines(k + 1).StartPoint, offsetLines(k + 1).EndPoint, erreur, True)
                offsetPoints.Add(point)
            Next
            Dim first = GetIntersection(offsetLines.Last.StartPoint, offsetLines.Last.EndPoint, offsetLines.First.StartPoint, offsetLines.First.EndPoint, erreur, True)
            offsetPoints.Add(first)
            Return New Surface(offsetPoints, Type)
        End If
    End Function

    ''' <summary>
    ''' Offset of Surface perimeter
    ''' </summary>
    ''' <param name="dist"> Distance of Offset </param>
    ''' <returns> New Surface </returns>
    Public Function Offset2(dist As Double) As Surface
        If Type = SurfaceType.Circle Then
            Dim dir = New Vector(Points(0), Points(1))
            If dist >= dir.Length Then
                Return Nothing
            Else
                Dim offsetpoints As New List(Of Point) From {
                Points(0),
                New Point(Points(1).X + dist * dir.ProperX, Points(1).Y + dist * dir.ProperY)
            }
                Return New Surface(offsetpoints, Type)
            End If
        ElseIf Type = SurfaceType.Ellipse Then
            If dist >= Math.Min(Distance(Points(0), Points(1)), Distance(Points(0), Points(2))) Then
                Return Nothing
            Else
                Dim offsetpoints As New List(Of Point) From {
                    Points(0),
                    New Vector(Points(0), Points(1)).GetPointAtDistance(dist, Points(1)),
                    New Vector(Points(0), Points(2)).GetPointAtDistance(dist, Points(2))
                }
                Return New Surface(offsetpoints, Type)
            End If
        Else
            Dim angles = GetInnerAngles().Item1
            Dim distances As New List(Of Double)
            Dim distances2 As New List(Of Double)
            Dim lines As New List(Of Line)

            lines.Add(New Line(Points.Last, Points.First))
            For i = 0 To Points.Count - 2
                lines.Add(New Line(Points(i), Points(i + 1)))
            Next

            For Each angle In angles
                distances.Add(dist / (Math.Sin(Math.Abs(angle / 180 * Math.PI) / 2)))
                distances2.Add(dist / (Math.Sqrt((1 - Math.Sqrt(1 - Math.Sin(angle / 180 * Math.PI) * Math.Sin(angle / 180 * Math.PI))) / 2)))
            Next

            Dim bissectrices As New List(Of Vector)
            For k = 0 To lines.Count - 2
                bissectrices.Add((lines(k).Vector.ScaleOneVector - lines(k + 1).Vector.ScaleOneVector).ScaleOneVector)
            Next
            bissectrices.Add((lines.Last.Vector.ScaleOneVector - lines.First.Vector.ScaleOneVector).ScaleOneVector)

            Dim newPoints As New List(Of Point)
            For k = 0 To Points.Count - 1
                newPoints.Add(New Point(Points(k).X + distances(k) * bissectrices(k).X, Points(k).Y + distances(k) * bissectrices(k).Y, Points(k).Z + distances(k) * bissectrices(k).Z))
            Next

            Return New Surface(newPoints, Type)
        End If
    End Function

    ''' <summary>
    ''' Clean duplicated points of the Surface, optional fractional precision
    ''' </summary>
    ''' <param name="prec"> Fractional precision for mrging two points </param>
    ''' <returns></returns>
    Public Function CleanPoints(Optional prec As Integer = 4) As Surface
        Dim newPoints As New List(Of Point) From {
            Points.First
        }
        If Type = SurfaceType.Circle Then
            If Distance(Points(1), Points(0)) > Math.Pow(10, -prec) Then
                newPoints.Add(Points(1))
            End If
            If newPoints.Count > 1 Then
                Return New Surface(newPoints, Type)
            Else
                Return Nothing
            End If
        ElseIf Type = SurfaceType.Ellipse Then
            If Distance(Points(1), Points(0)) > Math.Pow(10, -prec) Then
                newPoints.Add(Points(1))
            End If
            If Distance(Points(2), Points(0)) > Math.Pow(10, -prec) Then
                newPoints.Add(Points(2))
            End If
            If newPoints.Count > 2 Then
                Return New Surface(newPoints, Type)
            Else
                Return Nothing
            End If
        Else
            For k = 1 To Points.Count - 2
                If Distance(Points(k), newPoints.Last) > Math.Pow(10, -prec) Then
                    newPoints.Add(Points(k))
                End If
            Next
            If Distance(Points.Last, newPoints.Last) > Math.Pow(10, -prec) AndAlso Distance(Points.Last, newPoints.First) > Math.Pow(10, -prec) Then
                newPoints.Add(Points.Last)
            End If
            If newPoints.Count > 2 Then
                Return New Surface(newPoints, Type)
            Else
                Return Nothing
            End If
        End If
    End Function

    ''' <summary>
    ''' Fused nearby Aligned Edges.newEdges Egdes must be ordered.
    ''' </summary>
    ''' <param name="fusionAngle">Maximum angle in degree of fusion between 2 edges side by side, 0 is not advise.</param>
    ''' <param name="tolerance">max distance fusion between 2 point</param>
    ''' <returns></returns>
    Public Function CleanEdges(Optional fusionAngle As Double = 2, Optional tolerance As Double? = Nothing) As Surface
        Dim newEdges As New List(Of Curve)
        If Type = SurfaceType.Polygon OrElse Type = SurfaceType.Rectangle Then
            'Dim previousPrecision = Application.PRECISION
            'Application.PRECISION = 2
            Dim previousEdge = Edges.FirstOrDefault()
            For i As Integer = 1 To Edges.Count
                Dim currentEdge = Nothing
                If i = Edges.Count Then
                    'Last Edge may be fused with first edge
                    currentEdge = newEdges.FirstOrDefault()
                Else
                    currentEdge = Edges(i)
                End If

                'Try Cast lines
                Dim l1 = CType(previousEdge, Line)
                Dim l2 = CType(currentEdge, Line)
                If l1 Is Nothing OrElse l2 Is Nothing Then
                    'not lines. Next
                    previousEdge = currentEdge
                    Continue For
                End If

                'Get vectors
                'Dim v1 = l1.Vector.ScaleOneVector
                'Dim v2 = l2.Vector.ScaleOneVector
                Dim v1 = l1.Vector.ScaleOneVector
                Dim v2 = l2.Vector.ScaleOneVector
                Dim v1Array = New Object() {v1.X, v1.Y, v1.Z}
                Dim v2Array = New Object() {v2.X, v2.Y, v2.Z}

                If v1.X = 0 AndAlso v1.Y = 0 AndAlso v1.Z = 0 OrElse v2.X = 0 AndAlso v2.Y = 0 AndAlso v2.Z = 0 Then
                    Continue For
                End If

                'Are Edges aligned and same orientation?
                If Math.Abs(v2.Angle(v1, True)) <= fusionAngle AndAlso DotProduct(v1Array, v2Array) > 0 Then
                    'Wich point must be fused?
                    Dim fusedEdge = Nothing
                    If Point.AlmostEquals(previousEdge.EndPoint, currentEdge.StartPoint, tolerance) Then
                        '===>0===>
                        fusedEdge = New Line(previousEdge.StartPoint, currentEdge.EndPoint)
                        previousEdge = fusedEdge
                    ElseIf Point.AlmostEquals(previousEdge.StartPoint, currentEdge.EndPoint, tolerance) Then
                        '<===0<===
                        fusedEdge = New Line(currentEdge.StartPoint, previousEdge.EndPoint)
                        previousEdge = fusedEdge
                    Else
                        'Edges not continuous
                        Throw New Exception("Edges not continuous")
                    End If
                    If i = Edges.Count Then
                        'Fuse first edge and last edge
                        newEdges(0) = fusedEdge
                    End If
                Else
                    'No fusion
                    newEdges.Add(previousEdge)
                    previousEdge = currentEdge
                End If
            Next
            'Application.PRECISION = previousPrecision
            Return New Surface(newEdges, Type)
        Else
            Return Me
        End If
    End Function


    Public Overrides Function ToString() As String
        Dim typeStr = ""
        Select Case Type
            Case SurfaceType.Circle
                typeStr = "Circle"
            Case SurfaceType.Rectangle
                typeStr = "Rectangle"
            Case SurfaceType.Polygon
                typeStr = "Polygon"
            Case SurfaceType.Ellipse
                typeStr = "Ellipse"
        End Select
        Dim pointsToStr As String = ""
        For Each point In Points
            pointsToStr += point.ToString + ","
        Next
        Return typeStr + " : {" + pointsToStr.Substring(0, pointsToStr.Length - 1) + "}"
    End Function
End Class
