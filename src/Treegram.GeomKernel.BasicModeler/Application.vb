Imports System.Globalization
Imports System.Runtime.InteropServices
Imports PolyBoolCS

''' <summary>
''' Many geometrical functions that use different types of objects
''' </summary>
Public Module Application
    'prochainement séparation de calcul entre double et decimal, précision max double = RealPrecision, decimal = 28
    'Tous le système sera repenser pour demander pour certaines fonctions bien précise appelées la précision souhaitée
    'Suppression d'Approximate calculus et de la Precision globale

    <Obsolete>
    Public Precision As Integer = 3

    ''' <summary>
    ''' Maximum of significant digits for a number.
    ''' </summary>
    Public Const MaxPrecision = 14


    Private _realPrecision As Integer = 7
    ''' <summary>
    ''' Significant digits for real number.
    ''' </summary>
    Public Property RealPrecision As Integer
        Get
            Return _realPrecision
        End Get
        Set(value As Integer)
            If (value > MaxPrecision) Then
                Throw New ArgumentException($"Precision is higher than double precision. Max is {MaxPrecision}.")
            End If
            If (value <= 0) Then
                Throw New ArgumentException("Precision must be, at least, equals to 1.")
            End If
            _realPrecision = value
        End Set
    End Property

    ''' <summary>
    ''' Negative integer, ex : -2 = round to 10^-2
    ''' </summary>
    ''' <returns></returns>
    Public Property DefaultTolerance As Integer = -4

    ''' <summary>
    ''' Rounding of number to significant digits.
    ''' </summary>
    ''' <param name="d"></param>
    ''' <param name="digits"></param>
    ''' <returns></returns>
    Public Function RoundToSignificantDigits(ByVal d As Double, ByVal digits As Integer) As Double
        If d = 0 Then Return 0
        Dim negDig As Integer = Math.Floor(Math.Log10(Math.Abs(d))) + 1
        If negDig > digits OrElse negDig < (digits - 15) Then
            'Dim unPow = d * Math.Pow(10, digits - negDig)
            'Return Math.Round(unPow, 0) * Math.Pow(10, negDig - digits)

            Dim valuePowered = d * Math.Pow(10, digits - negDig)
            Dim significativeResult = Math.Round(valuePowered, 0)
            significativeResult *= Math.Pow(10, negDig - digits)

            Dim textualResult = significativeResult.ToString(CultureInfo.InvariantCulture)
            Dim textualResultContainsSeparator = (textualResult.Contains("."))

            Dim size = digits

            If textualResultContainsSeparator Then
                size += 1
            End If

            If d < 0 Then
                size += 1
            End If

            If textualResult.Length <= size Then Return significativeResult

            Dim scientificResult = significativeResult.ToString("E", CultureInfo.InvariantCulture)
            Dim split = scientificResult.Split("E")
            Dim scientificContainsSeparator = (split(0).Contains("."))

            If split(0).Length > size Then
                significativeResult = Double.Parse($"{split(0).Substring(0, size)}E{split(1)}", CultureInfo.InvariantCulture)
            Else
                significativeResult = Double.Parse(scientificResult, CultureInfo.InvariantCulture)
            End If

            Return significativeResult
        Else
            Return Math.Round(d, digits - negDig)
        End If
    End Function


#Region "Distances"
    ''' <summary>
    ''' Distance 3D point to point with
    ''' </summary>
    ''' Trustable function, only depends on Windows class
    ''' <param name="pointA"> First Point </param>
    ''' <param name="pointB"> Second Point </param>
    ''' <returns> Distance between points </returns>
    Public Function Distance(pointA As Point, pointB As Point) As Double
        Return Math.Round(RoundToSignificantDigits(Math.Sqrt((pointB.X - pointA.X) * (pointB.X - pointA.X) + (pointB.Y - pointA.Y) * (pointB.Y - pointA.Y) + (pointB.Z - pointA.Z) * (pointB.Z - pointA.Z)), RealPrecision), -DefaultTolerance)
    End Function


    ''' <summary>
    ''' Distance 2D Vector to Vector (if infiniteLines = true, vector = line else vector = segment).
    ''' Works ONLY IN 2D for now (parameters X and Y)!
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="infiniteLines"> Condition of evaluation in segment or line </param>
    ''' <returns> The Distance between to vectors (as segments or lines) </returns>
    <Obsolete("Old function for vectors works only in 2D")>
    Public Function OldDistance(v1 As Vector, v2 As Vector, Optional infiniteLines As Boolean = True) As Double
        If v1.Origin.Z <> v2.Origin.Z OrElse v1.Z <> 0 OrElse v2.Z <> 0 Then
            Throw New ArgumentException("Vectors have to be on the same horizontal plane (Z=constant)")
        End If
        Dim result As Double = Double.PositiveInfinity
        Dim saveZ = v1.Origin.Z
        v1.Origin.Z = 0
        v2.Origin.Z = 0
        'v1.EndPoint.Z = 0
        'v2.Z = 0
        Dim s1 = v1.ScaleOneVector
        Dim s2 = v2.ScaleOneVector

        If s1 = s2 Or s1 = s2.Reverse Then ' If vectors share the same direction
            If infiniteLines Then
                Dim normalPoint = GetNormalPoint(v1.Origin, v2.Origin, v2.EndPoint)
                result = Distance(v1.Origin, normalPoint)
            Else
                Dim normalOriginv1Onv2 = GetNormalPoint(v1.Origin, v2.Origin, v2.EndPoint),
                    normalEndPointv1Onv2 = GetNormalPoint(v1.EndPoint, v2.Origin, v2.EndPoint),
                    normalOriginv2Onv1 = GetNormalPoint(v2.Origin, v1.Origin, v1.EndPoint),
                    normalEndPointv2Onv1 = GetNormalPoint(v2.EndPoint, v1.Origin, v1.EndPoint)
                If Distance(normalOriginv1Onv2, v2.Origin, v2.EndPoint, infiniteLines) = 0.0 Then
                    result = Distance(v1.Origin, normalOriginv1Onv2)
                ElseIf Distance(normalEndPointv1Onv2, v2.Origin, v2.EndPoint, infiniteLines) = 0.0 Then
                    result = Distance(v1.EndPoint, normalEndPointv1Onv2)
                ElseIf Distance(normalOriginv2Onv1, v1.Origin, v1.EndPoint, infiniteLines) = 0.0 Then
                    result = Distance(v2.Origin, normalOriginv2Onv1)
                ElseIf Distance(normalEndPointv2Onv1, v1.Origin, v1.EndPoint, infiniteLines) = 0.0 Then
                    result = Distance(v2.EndPoint, normalEndPointv2Onv1)
                Else
                    Dim ov1Ov2 = Distance(v1.Origin, v2.Origin)
                    Dim ev1Ov2 = Distance(v1.EndPoint, v2.Origin)
                    Dim ov1Ev2 = Distance(v1.Origin, v2.EndPoint)
                    Dim ev1Ev2 = Distance(v1.EndPoint, v2.EndPoint)
                    result = {ov1Ov2, ev1Ov2, ov1Ev2, ev1Ev2}.Min
                End If
            End If
        ElseIf infiniteLines Then ' Infinite 2D lines that ain't colinear always intersect
            result = 0.0
        Else
            ' STEP 1 : Check if lines intersect
            Dim doIntersect As Boolean = False
            Dim errorCode As ErrorCode = ErrorCode.UndefinedError
            Dim intersection = GetIntersection(v1.Origin, v1.EndPoint, v2.Origin, v2.EndPoint, errorCode, True)

            If errorCode = ErrorCode.NoError Then
                If Distance(intersection, v1.Origin, v1.EndPoint, infiniteLines) = 0 And
                     Distance(intersection, v2.Origin, v2.EndPoint, infiniteLines) = 0 Then
                    doIntersect = True
                    result = 0.0
                End If
            End If

            If Not doIntersect Then
                ' STEP 2 : Check if projection of the extremities of each line intersect with the other vector
                ' If so, return the lowest distance 
                Dim nv1OV2 = GetNormalPoint(v1.Origin, v2.Origin, v2.EndPoint)
                Dim dNv1OV2 = Distance(nv1OV2, v2, infiniteLines)
                Dim iNv1OV2 As Boolean = dNv1OV2 = 0

                Dim nv1EV2 = GetNormalPoint(v1.EndPoint, v2.Origin, v2.EndPoint)
                Dim dNv1EV2 = Distance(nv1EV2, v2, infiniteLines)
                Dim iNv1EV2 As Boolean = dNv1EV2 = 0

                Dim nv2OV1 = GetNormalPoint(v2.Origin, v1.Origin, v1.EndPoint)
                Dim dNv2OV1 = Distance(nv2OV1, v1, infiniteLines)
                Dim iNv2OV1 As Boolean = dNv2OV1 = 0

                Dim nv2EV1 = GetNormalPoint(v2.EndPoint, v1.Origin, v1.EndPoint)
                Dim dNv2EV1 = Distance(nv2EV1, v1, infiniteLines)
                Dim iNv2EV1 As Boolean = dNv2EV1 = 0

                If iNv1OV2 Or iNv1EV2 Or iNv2OV1 Or iNv2EV1 Then
                    result = {If(iNv1OV2, Distance(v1.Origin, v2), Double.PositiveInfinity),
                                If(iNv1EV2, Distance(v1.EndPoint, v2), Double.PositiveInfinity),
                                If(iNv2OV1, Distance(v2.Origin, v1), Double.PositiveInfinity),
                                If(iNv2EV1, Distance(v2.EndPoint, v1), Double.PositiveInfinity)}.Min
                Else
                    ' STEP 3 : Get the distance between all extremities and return the lowest
                    Dim dV1OV2O = Distance(v1.Origin, v2.Origin)
                    Dim dV1OV2E = Distance(v1.Origin, v2.EndPoint)
                    Dim dV1EV2O = Distance(v1.EndPoint, v2.Origin)
                    Dim dV1EV2E = Distance(v1.EndPoint, v2.EndPoint)
                    result = {dV1OV2O, dV1OV2E, dV1EV2O, dV1EV2E}.Min
                End If
            End If

        End If
        Return result
    End Function

    ''' <summary>
    ''' Distance 3D Vector to Vector (if infiniteLines = true, vector = line else vector = segment).
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="infinitelines"> Condition of evaluation in segment or line </param>
    ''' <returns> The Distance between to vectors (as segments or lines) </returns>
    Public Function Distance(v1 As Vector, v2 As Vector, Optional infinitelines As Boolean = True) As Double
        Dim r = v1 * v2
        Dim d1 = v1.InternLength * v1.InternLength
        Dim d2 = v2.InternLength * v2.InternLength
        Dim s1 = v1.X * (v2.Origin.X - v1.Origin.X) + v1.Y * (v2.Origin.Y - v1.Origin.Y) + v1.Z * (v2.Origin.Z - v1.Origin.Z)
        Dim s2 = v2.X * (v2.Origin.X - v1.Origin.X) + v2.Y * (v2.Origin.Y - v1.Origin.Y) + v2.Z * (v2.Origin.Z - v1.Origin.Z)
        Dim t = 0.0
        Dim u = 0.0
Etape1:
        If d1 = 0 Xor d2 = 0 Then
            GoTo Etape4
        ElseIf d1 = 0 Then
            GoTo Etape5
        ElseIf d1 * d2 - r * r = 0 Then
            GoTo Etape3
        End If
Etape2:
        t = (s1 * d2 - s2 * r) / (d1 * d2 - r * r)
        If Not infinitelines Then
            t = Math.Max(Math.Min(t, 1), 0)
        End If
Etape3:
        u = (t * r - s2) / d2
        If Not infinitelines AndAlso (u < 0 OrElse u > 1) Then
            u = Math.Max(Math.Min(u, 1), 0)
        Else
            GoTo Etape5
        End If
Etape4:
        t = (u * r + s1) / d1
        If Not infinitelines Then
            t = Math.Max(Math.Min(t, 1), 0)
        End If
Etape5:
        Dim dd = Math.Pow(v1.X * t - v2.X * u - (v2.Origin.X - v1.Origin.X), 2) + Math.Pow(v1.Y * t - v2.Y * u - (v2.Origin.Y - v1.Origin.Y), 2) + Math.Pow(v1.Z * t - v2.Z * u - (v2.Origin.Z - v1.Origin.Z), 2)
        Return Math.Round(RoundToSignificantDigits(Math.Sqrt(dd), RealPrecision), -DefaultTolerance)
    End Function

    ''' <summary>
    ''' Distance between segments with return of parameters to have closest points
    ''' </summary>
    ''' <param name="v1"></param>
    ''' <param name="v2"></param>
    ''' <param name="t"></param>
    ''' <param name="u"></param>
    ''' <returns></returns>
    Private Function Distance(v1 As Vector, v2 As Vector, ByRef t As Double, ByRef u As Double) As Double
        Dim r = v1 * v2
        Dim d1 = v1.InternLength * v1.InternLength
        Dim d2 = v2.InternLength * v2.InternLength
        Dim s1 = v1.X * (v2.Origin.X - v1.Origin.X) + v1.Y * (v2.Origin.Y - v1.Origin.Y) + v1.Z * (v2.Origin.Z - v1.Origin.Z)
        Dim s2 = v2.X * (v2.Origin.X - v1.Origin.X) + v2.Y * (v2.Origin.Y - v1.Origin.Y) + v2.Z * (v2.Origin.Z - v1.Origin.Z)
        t = 0.0
        u = 0.0
Etape1:
        If d1 = 0 Xor d2 = 0 Then
            GoTo Etape4
        ElseIf d1 = 0 Then
            GoTo Etape5
        ElseIf d1 * d2 - r * r = 0 Then
            GoTo Etape3
        End If
Etape2:
        t = (s1 * d2 - s2 * r) / (d1 * d2 - r * r)
        t = Math.Max(Math.Min(t, 1), 0)
Etape3:
        u = (t * r - s2) / d2
        If (u < 0 OrElse u > 1) Then
            u = Math.Max(Math.Min(u, 1), 0)
        Else
            GoTo Etape5
        End If
Etape4:
        t = (u * r + s1) / d1
        t = Math.Max(Math.Min(t, 1), 0)
Etape5:
        Dim dd = Math.Pow(v1.X * t - v2.X * u - (v2.Origin.X - v1.Origin.X), 2) + Math.Pow(v1.Y * t - v2.Y * u - (v2.Origin.Y - v1.Origin.Y), 2) + Math.Pow(v1.Z * t - v2.Z * u - (v2.Origin.Z - v1.Origin.Z), 2)
        Return Math.Round(RoundToSignificantDigits(Math.Sqrt(dd), RealPrecision), -DefaultTolerance)
    End Function

    ''' <summary>
    ''' Distance 3D Line to Line (if infiniteLines = true, line infinite else line = segment).
    ''' </summary>
    ''' <param name="l1"> First Line </param>
    ''' <param name="l2"> Second Line </param>
    ''' <param name="infiniteLines"> Condition of evaluation in segment or infinite line </param>
    ''' <returns> The Distance between to lines or segments. </returns>
    Public Function Distance(l1 As Line, l2 As Line, Optional infiniteLines As Boolean = True) As Double
        'If l1.Vector.Origin.Z <> l2.Vector.Origin.Z OrElse l1.Vector.Z <> 0 OrElse l2.Vector.Z <> 0 Then
        '    Throw New ArgumentException("Lines have to be on the same horizontal plane (Z=constant)")
        'End If
        Return Distance(l1.Vector, l2.Vector, infiniteLines)
    End Function

    ''' <summary>
    ''' Distance 3D Point to Vector (if infiniteLine = true, vector = line else vector = segment)
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="v"> Vector </param>
    ''' <param name="infiniteLine"> Condition of evaluation in segment or line </param>
    ''' <returns> Distance between Point and Vector (segment or line)</returns>
    Public Function Distance(p As Point, v As Vector, Optional infiniteLine As Boolean = True) As Double
        Return Distance(p, v.Origin, v.EndPoint, infiniteLine)
    End Function

    ''' <summary>
    ''' Distance 3D Point to Line (if infiniteLine = true, Line is infinite else line = segment)
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="l"> Line </param>
    ''' <param name="infiniteLine"> Condition of evaluation in segment or infinite line </param>
    ''' <returns></returns>
    Public Function Distance(p As Point, l As Line, Optional infiniteLine As Boolean = True) As Double
        Return Distance(p, l.StartPoint, l.EndPoint, infiniteLine)
    End Function

    ''' <summary>
    ''' Distance 3D Point P to Vectoor(P1,P2) (if infiniteLines = true, vector = line else vector = segment)
    ''' </summary> 
    ''' <param name="p"> Point to evaluate </param>
    ''' <param name="p1"> First Point for Vector definiiton </param>
    ''' <param name="p2"> Second Point for Vector definition </param>
    ''' <param name="infiniteLine"> Condition of evaluation in segment or line </param>
    ''' <returns> Distance between Point to Vector(P1,P2) </returns>
    Public Function Distance(p As Point, p1 As Point, p2 As Point, Optional infiniteLine As Boolean = True) As Double
        Dim normalPoint = GetNormalPoint(p, p1, p2)
        Dim length = Distance(p1, p2)
        Dim distanceP1 = Distance(normalPoint, p1)
        Dim distanceP2 = Distance(normalPoint, p2)

        If infiniteLine Then
            Return Distance(normalPoint, p)
        Else
            If (distanceP1 <= length And distanceP2 <= length) Then
                Return Distance(normalPoint, p)
            ElseIf distanceP1 < distanceP2 Then
                Return Distance(p, p1)
            Else
                Return Distance(p, p2)
            End If
        End If
    End Function

    ''' <summary>
    ''' Distance 2D Point to Surface
    ''' </summary>
    ''' Works ONLY IN 2D for now (parameters X and Y)!
    ''' Only with circles, rectangles and polygons (uncrossed)
    ''' <param name="pt"> Point </param>
    ''' <param name="surf"> Surface </param>
    ''' <returns> Distance between Point and Surface </returns>
    Public Function Distance(pt As Point, surf As Surface) As Double
        If surf.Points.Count < 2 Then
            Throw New Exception("Surf needs at least 3 points.")
        End If
        If Not (surf.WindingNumber(pt) Mod 2 = 1) Then
            If surf.Type = Surface.SurfaceType.Circle Then
                Return Distance(surf.Points(0), pt) - Distance(surf.Points(1), surf.Points(0))
            ElseIf surf.Type = Surface.SurfaceType.Rectangle OrElse surf.Type = Surface.SurfaceType.Polygon Then
                Dim dist = Double.PositiveInfinity
                For i = 0 To surf.Points.Count - 2
                    Dim line = New Line(surf.Points(i), surf.Points(i + 1))
                    If Distance(pt, line.Vector, False) < dist Then
                        dist = Distance(pt, line.Vector, False)
                    End If
                Next
                If Distance(pt, surf.Points.Last, surf.Points.First, False) < dist Then
                    dist = Distance(pt, surf.Points.Last, surf.Points.First, False)
                End If
                Return dist
            End If
        End If
        Return 0.0
    End Function

    ''' <summary>
    ''' Distance 2D Line to Surface
    ''' </summary>
    ''' Works ONLY IN 2D for now (parameters X and Y)!
    ''' Only with Circles, Rectangles and Polygons (uncrossed)
    ''' <param name="myLine"> Line </param>
    ''' <param name="surf"> Surface </param>
    ''' <returns> Distance between Line and Surface </returns>
    Public Function Distance(myLine As Line, surf As Surface) As Double
        If surf.Points.Count < 2 Then
            Throw New Exception("Surf needs at least 3 points.")
        End If
        Dim startPt = myLine.StartPoint
        Dim endPt = myLine.EndPoint
        If Not (surf.WindingNumber(startPt) Mod 2 = 1 OrElse surf.WindingNumber(endPt)) Then
            Dim dist = Double.PositiveInfinity
            If surf.Type = Surface.SurfaceType.Circle Then
                dist = Distance(surf.Points(0), myLine.Vector, False) - Distance(surf.Points(1), surf.Points(0))
                Return Math.Max(0.0, dist)
            ElseIf surf.Type = Surface.SurfaceType.Rectangle OrElse surf.Type = Surface.SurfaceType.Polygon Then
                Dim erreur As ErrorCode
                For i = 0 To surf.Points.Count - 2
                    Dim line = New Line(surf.Points(i), surf.Points(i + 1))
                    If Not IsNothing(GetIntersection(startPt, endPt, line.StartPoint, line.EndPoint, erreur, False)) Then
                        Return 0.0
                    End If
                    If Distance(myLine, line, False) < dist Then
                        dist = Distance(myLine, line, False)
                    End If
                Next

                If Not IsNothing(GetIntersection(startPt, endPt, surf.Points.Last, surf.Points.First, erreur, False)) Then
                    Return 0.0
                End If
                If Distance(myLine, New Line(surf.Points.Last, surf.Points.First), False) < dist Then
                    dist = Distance(myLine, New Line(surf.Points.Last, surf.Points.First), False)
                End If
                Return dist
            End If
        End If
        Return 0.0
    End Function

    ''' <summary>
    ''' Distance 2D Surface to Surface
    ''' </summary>
    ''' Works ONLY IN 2D for now (parameters X and Y)!
    ''' Only with Circles, Rectangles and Polygons (uncrossed)
    ''' <param name="surf1"> First Surface </param>
    ''' <param name="surf2"> Second Surface </param>
    ''' <returns> Distance between two Surfaces </returns>
    Public Function Distance(surf1 As Surface, surf2 As Surface) As Double
        If Not DoSurfacesIntersect(surf1, surf2) Then
            If surf1.Type = Surface.SurfaceType.Circle Xor surf2.Type = Surface.SurfaceType.Circle Then
                If surf1.Type = Surface.SurfaceType.Circle Then
                    Dim dist = Double.PositiveInfinity
                    For k = 0 To surf2.Points.Count - 2
                        Dim line = New Line(surf2.Points(k), surf2.Points(k + 1))
                        dist = Math.Min(dist, Distance(line, surf1))
                        If dist = 0 Then Exit For
                    Next
                    Dim lastLine = New Line(surf2.Points.Last, surf2.Points.First)
                    dist = Math.Min(dist, Distance(lastLine, surf1))
                    Return dist
                Else
                    Dim dist = Double.PositiveInfinity
                    For k = 0 To surf1.Points.Count - 2
                        Dim line = New Line(surf1.Points(k), surf1.Points(k + 1))
                        dist = Math.Min(dist, Distance(line, surf2))
                        If dist = 0 Then Exit For
                    Next
                    Dim lastLine = New Line(surf1.Points.Last, surf1.Points.First)
                    dist = Math.Min(dist, Distance(lastLine, surf2))
                    Return dist
                End If
            Else
                If surf1.Type = Surface.SurfaceType.Circle Then
                    Return Distance(surf1.Points(0), surf2.Points(0)) - Distance(surf1.Points(0), surf1.Points(1)) - Distance(surf2.Points(0), surf2.Points(1))
                ElseIf (surf1.Type = Surface.SurfaceType.Rectangle OrElse surf1.Type = Surface.SurfaceType.Polygon) AndAlso surf2.Type = Surface.SurfaceType.Rectangle OrElse surf2.Type = Surface.SurfaceType.Polygon Then
                    Dim lines As New List(Of Line)
                    For k = 0 To surf1.Points.Count - 2
                        lines.Add(New Line(surf1.Points(k), surf1.Points(k + 1)))
                    Next
                    lines.Add(New Line(surf1.Points.Last, surf1.Points.First))
                    Dim dist = Double.PositiveInfinity
                    For Each line In lines
                        dist = Math.Min(Distance(line, surf2), dist)
                        If dist = 0 Then Exit For
                    Next
                    Return dist
                End If
            End If
        End If
        Return 0.0
    End Function

    ''' <summary>
    ''' Distance 3D between point and plane
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="plane"></param>
    ''' <returns></returns>
    Public Function Distance(point As Point, plane As Plane)
        If plane.Vector.InternLength = 0 Then
            Throw New ArgumentNullException("Plane's vector is null.")
        End If
        Return Math.Round(Math.Abs(plane.Vector * New Vector(plane.Origin, point)) / plane.Vector.InternLength, RealPrecision)
    End Function

    ''' <summary>
    ''' Distance 3D between point and triangle.
    ''' </summary>
    ''' <param name="point"></param>
    ''' <param name="triangle"></param>
    ''' <returns></returns>
    Public Function Distance(point As Point, triangle As Triangle) As Double
        Dim a = triangle.P1
        Dim b = triangle.P2
        Dim c = triangle.P3
        Dim plan As Plane = Nothing
        Try
            plan = New Plane(a, b, c)
        Catch ex As ArgumentException
            Return {Distance(point, New Vector(a, b), False), Distance(point, New Vector(b, c), False), Distance(point, New Vector(c, a), False)}.Min
        End Try

        Dim proj = plan.Projection(point)
        Dim vectAb = New Vector(a, b)
        Dim vectAc = New Vector(a, c)
        Dim det = vectAb.X * vectAc.Y * plan.Vector.Z + vectAb.Y * vectAc.Z * plan.Vector.X + vectAb.Z * vectAc.X * plan.Vector.Y - vectAb.Z * vectAc.Y * plan.Vector.X - vectAb.X * vectAc.Z * plan.Vector.Y - vectAb.Y * vectAc.X * plan.Vector.Z
        Dim up2 = (vectAc.Y * plan.Vector.Z - vectAc.Z * plan.Vector.Y) / det
        Dim uq2 = -(vectAc.X * plan.Vector.Z - vectAc.Z * plan.Vector.X) / det
        Dim ur2 = vectAc.X * plan.Vector.Y - vectAc.Y * plan.Vector.X / det
        Dim vp2 = -(vectAb.Y * plan.Vector.Z - vectAb.Z * plan.Vector.Y) / det
        Dim vq2 = (vectAb.X * plan.Vector.Z - vectAb.Z * plan.Vector.X) / det
        Dim vr2 = -(vectAb.X * plan.Vector.Y - vectAb.Y * plan.Vector.X) / det
        Dim wp2 = (vectAb.Y * vectAc.Z - vectAb.Z * vectAc.Y) / det
        Dim wq2 = -(vectAb.X * vectAc.Z - vectAb.Z * vectAc.X) / det
        Dim wr2 = (vectAb.X * vectAc.Y - vectAb.Y * vectAc.X) / det
        Dim matrice As Double(,) = {{up2, vp2, wp2, 0}, {uq2, vq2, wq2, 0}, {ur2, vr2, wr2, 0}, {0, 0, 0, 1}}
        Dim chgProj = TransformMat4(proj, matrice)
        Dim chgA = TransformMat4(a, matrice)
        Dim chgB = TransformMat4(b, matrice)
        Dim chgC = TransformMat4(c, matrice)
        Dim alpha = Vector.DeterminantZ(New Vector(chgProj, chgB), New Vector(chgProj, chgC))
        Dim beta = Vector.DeterminantZ(New Vector(chgProj, chgC), New Vector(chgProj, chgA))
        Dim gamma = Vector.DeterminantZ(New Vector(chgProj, chgA), New Vector(chgProj, chgB))

        If alpha <= 0 Then
            If beta <= 0 Then
                Return Distance(point, c)
            ElseIf gamma <= 0 Then
                Return Distance(point, b)
            Else
                Return Distance(point, b, c, False)
            End If
        Else
            If beta <= 0 Then
                If gamma <= 0 Then
                    Return Distance(point, a)
                Else
                    Return Distance(point, a, c, False)
                End If
            ElseIf gamma < 0 Then
                Return Distance(point, a, b, False)
            Else
                Return Distance(point, proj)
            End If
        End If
    End Function

    ''' <summary>
    ''' Distance 3D triangle to triangle.
    ''' </summary>
    ''' <param name="triangle1"></param>
    ''' <param name="triangle2"></param>
    ''' <returns></returns>
    Public Function Distance(triangle1 As Triangle, triangle2 As Triangle) As Double
        Dim sortedDistances = New SortedDictionary(Of Double, Tuple(Of Point, Point, Integer, Integer))
        Dim vectors1 = New List(Of Vector) From {
            New Vector(triangle1.P1, triangle1.P2),
            New Vector(triangle1.P2, triangle1.P3),
            New Vector(triangle1.P3, triangle1.P1)
        }
        Dim vectors2 = New List(Of Vector) From {
            New Vector(triangle2.P1, triangle2.P2),
            New Vector(triangle2.P2, triangle2.P3),
            New Vector(triangle2.P3, triangle2.P1)
        }

        For i = 0 To 2
            For j = 0 To 2
                Dim vect1 = vectors1(i)
                Dim vect2 = vectors2(j)
                Dim t, u As Double
                Dim dd = Distance(vect1, vect2, t, u)
                Try
                    sortedDistances.Add(dd, New Tuple(Of Point, Point, Integer, Integer)(vect1.GetPointAtParameter(t), vect2.GetPointAtParameter(u), (i + 2) Mod 3, (j + 2) Mod 3))
                Catch
                End Try
            Next
        Next
        Dim dist = sortedDistances.Keys(0)
        If dist = 0 Then
            Return dist
        End If
        Dim pt1 = sortedDistances.Values(0).Item1
        Dim pt2 = sortedDistances.Values(0).Item2
        Dim norm1 = New Vector(pt1, pt2)
        Dim norm2 = New Vector(pt2, pt1)

        'Cas classique où la distance est le minimum des distances côtés à côtés
        If New Vector(pt1, triangle1.Points(sortedDistances.Values(0).Item3)) * norm1 < 0 AndAlso New Vector(pt2, triangle2.Points(sortedDistances.Values(0).Item4)) * norm2 < 0 Then
            Return dist
        End If

        Dim plane1 As Plane = Nothing
        Dim plane2 As Plane = Nothing
        Try
            plane1 = New Plane(triangle1.P1, triangle1.P2, triangle1.P3)
        Catch ex As ArgumentException
        End Try
        Try
            plane2 = New Plane(triangle2.P1, triangle2.P2, triangle2.P3)
        Catch ex As ArgumentException
        End Try



        Dim pp1 As Point
        Dim pp2 As Point
        Dim pp3 As Point
        If plane2 Is Nothing Then
            GoTo suiv
        End If
        Dim first = (New Vector(plane2.Origin, triangle1.P1) * triangle2.Normal > 0)
suiv:
        If plane2 IsNot Nothing AndAlso (New Vector(plane2.Origin, triangle1.P2) * triangle2.Normal > 0) = first AndAlso (New Vector(plane2.Origin, triangle1.P3) * triangle2.Normal > 0) = first Then
            'Cas où on projette les points du triangle 1 sur le plan du 2
            pp1 = plane2.Projection(triangle1.P1)
            pp2 = plane2.Projection(triangle1.P2)
            pp3 = plane2.Projection(triangle1.P3)
            Dim sortedDist = New SortedDictionary(Of Double, Point)
            Try
                sortedDist.Add(Distance(triangle1.P1, pp1), pp1)
            Catch
            End Try
            Try
                sortedDist.Add(Distance(triangle1.P2, pp2), pp2)
            Catch
            End Try
            Try
                sortedDist.Add(Distance(triangle1.P3, pp3), pp3)
            Catch
            End Try

            If Distance(sortedDist.Values(0), triangle2) = 0 Then
                'Cas de projection
                Return sortedDist.Keys(0)
            Else
                'Cas dégénéré
                Return dist
            End If
        ElseIf plane1 IsNot Nothing Then
            first = (New Vector(plane1.Origin, triangle2.P1) * triangle1.Normal > 0)
            If (New Vector(plane1.Origin, triangle2.P2) * triangle1.Normal > 0) = first AndAlso (New Vector(plane1.Origin, triangle2.P3) * triangle1.Normal > 0) = first Then
                ' Cas où on projette les points du triagle 2 sur le plan du 1
                pp1 = plane1.Projection(triangle2.P1)
                pp2 = plane1.Projection(triangle2.P2)
                pp3 = plane1.Projection(triangle2.P3)
                Dim sortedDist = New SortedDictionary(Of Double, Point)
                Try
                    sortedDist.Add(Distance(triangle2.P1, pp1), pp1)
                Catch
                End Try
                Try
                    sortedDist.Add(Distance(triangle2.P2, pp2), pp2)
                Catch
                End Try
                Try
                    sortedDist.Add(Distance(triangle2.P3, pp3), pp3)
                Catch
                End Try

                If Distance(sortedDist.Values(0), triangle1) = 0 Then
                    'Cas de projection
                    Return sortedDist.Keys(0)
                Else
                    'Cas dégénéré
                    Return dist
                End If
            Else
                'Cas d'intersection
                Return 0.0
            End If
        End If

        Return dist
    End Function

    ''' <summary>
    ''' Distance 3D between vector and triangle.
    ''' </summary>
    ''' <param name="vector"></param>
    ''' <param name="triangle"></param>
    ''' <returns></returns>
    Public Function Distance(vector As Vector, triangle As Triangle) As Double
        Dim sortedDistances = New SortedDictionary(Of Double, Tuple(Of Point, Point, Integer))
        Dim vectors = New List(Of Vector) From {
            New Vector(triangle.P1, triangle.P2),
            New Vector(triangle.P2, triangle.P3),
            New Vector(triangle.P3, triangle.P1)
        }
        For i = 0 To 2
            Dim vect = vectors(i)
            Dim t, u As Double
            Dim dd = Distance(vect, vector, t, u)
            Try
                sortedDistances.Add(dd, New Tuple(Of Point, Point, Integer)(vector.GetPointAtParameter(t), vect.GetPointAtParameter(u), (i + 2) Mod 3))
            Catch
            End Try
        Next
        Dim dist = sortedDistances.Keys(0)
        If dist = 0 Then
            Return dist
        End If
        Dim pt1 = sortedDistances.Values(0).Item1
        Dim pt2 = sortedDistances.Values(0).Item2
        Dim norm2 = New Vector(pt2, pt1)

        'Cas classique où la distance est le minimum des distances côtés à côtés
        If New Vector(pt2, triangle.Points(sortedDistances.Values(0).Item3)) * norm2 < 0 Then
            Return dist
        End If

        Dim plane As Plane = Nothing
        Try
            plane = New Plane(triangle.P1, triangle.P2, triangle.P3)
        Catch ex As ArgumentException
        End Try

        Dim pp1 As Point
        Dim pp2 As Point

        If plane IsNot Nothing AndAlso (New Vector(plane.Origin, vector.Origin) * triangle.Normal > 0) = (New Vector(plane.Origin, vector.EndPoint) * triangle.Normal > 0) Then
            'cas de projection
            pp1 = plane.Projection(vector.Origin)
            pp2 = plane.Projection(vector.EndPoint)
            Dim sortedDist = New SortedDictionary(Of Double, Point)
            Try
                sortedDist.Add(Distance(vector.Origin, pp1), pp1)
            Catch
            End Try
            Try
                sortedDist.Add(Distance(vector.EndPoint, pp2), pp2)
            Catch
            End Try
            If Distance(sortedDist.Values(0), triangle) = 0 Then
                'Cas de projection
                Return sortedDist.Keys(0)
            Else
                'Cas dégénéré
                Return dist
            End If
        Else
            'Cas intersection
            Return 0
        End If
        Return dist
    End Function

    ''' <summary>
    ''' Distance 3D between line and triangle
    ''' </summary>
    ''' <param name="line"></param>
    ''' <param name="triangle"></param>
    ''' <returns></returns>
    Public Function Distance(line As Line, triangle As Triangle) As Double
        Return Distance(line.Vector, triangle)
    End Function
#End Region


    ''' <summary>
    ''' Check in 3D if the vectors = segments are on the exact same line
    ''' </summary>
    ''' It's not a check of parallelism
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance">Set the Tolerance as 10^-tolerance (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
    ''' <returns> Result of check as Boolean </returns>
    <Obsolete("Use the function Vector.AreAligned")>
    Public Function AreVectorsOnTheSameLine(v1 As Vector, v2 As Vector, Optional tolerance As Integer? = Nothing) As Boolean ''A changer par une tolerance
        If tolerance Is Nothing Then tolerance = -DefaultTolerance
        If Vector.AreVectorsParallel(v1, v2, tolerance) Then
            If Distance(v1.Origin, v2, True) <= 1 * Math.Pow(10, -tolerance) AndAlso Distance(v1.EndPoint, v2, True) <= 1 * Math.Pow(10, -tolerance) Then
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Return the middle point of the segment(pointA, pointB)
    ''' </summary>
    ''' Trustable function, only depends on Windows class
    ''' <param name="pointA"> First Point </param>
    ''' <param name="pointB"> Second Point </param>
    ''' <returns> Middle Point of segment(pointA,pointB) </returns>
    Public Function MidPoint(pointA As Point, pointB As Point) As Point
        Return New Point((pointA.X + pointB.X) / 2, (pointA.Y + pointB.Y) / 2, (pointA.Z + pointB.Z) / 2)
    End Function

    Public Enum ErrorCode
        NoError
        ParallelLines
        CoincidentLines
        NoLineDetected
        UndefinedError
        ParallelToPlan
        NoPlanDetected
    End Enum

    ''' <summary>
    ''' Orthogonal projection in 3D of point P on Line(P1,P2)
    ''' </summary>
    ''' <param name="p"> Point to project </param>
    ''' <param name="p1"> First Point for Line definition </param>
    ''' <param name="p2"> Second Point for Line definition </param>
    ''' <returns> Projected Point </returns>
    Public Function GetNormalPoint(p As Point, p1 As Point, p2 As Point) As Point
        If p1 = p2 Then
            Throw New ArgumentNullException("Line is null, P1=P2")
        Else
            Dim k As Double = New Vector(p1, p) * New Vector(p1, p2) / (New Vector(p1, p2).InternLength * New Vector(p1, p2).InternLength)
            Dim x4 As Double = p1.X + k * (p2.X - p1.X)
            Dim y4 As Double = p1.Y + k * (p2.Y - p1.Y)
            Dim z4 As Double = p1.Z + k * (p2.Z - p1.Z)
            Return New Point() With {.X = x4, .Y = y4, .Z = z4}
        End If
    End Function

#Region "Reliquat du passé"
    'Public Function GetIntersection(lineA_1 As System.Drawing.Point, lineA_2 As System.Drawing.Point, lineB_1 As System.Drawing.Point, lineB_2 As System.Drawing.Point) As System.Drawing.Point
    '    Dim ec As ErrorCode
    '    Dim result = GetIntersection(lineA_1, lineA_2, lineB_1, lineB_2, ec)
    '    If ec = ErrorCode.noError Then
    '        Return result
    '    Else
    '        Throw New Exception
    '    End If
    'End Function
#End Region

    ''' <summary>
    ''' Get the Intersection 2D Point of two segments or infinite lines
    ''' </summary>
    ''' <param name="lineA1"> First Point of first Line </param>
    ''' <param name="lineA2"> Second Point of first Line </param>
    ''' <param name="lineB1"> First Point of second Line </param>
    ''' <param name="lineB2"> Second Point of second Line </param>
    ''' <param name="errorCode"> Error </param>
    ''' <param name="infiniteLines"> Infinite Lines or Segments </param>
    ''' <returns> Intersection Point </returns>
    Public Function GetIntersection(lineA1 As Point, lineA2 As Point, lineB1 As Point, lineB2 As Point, ByRef errorCode As ErrorCode, Optional infiniteLines As Boolean = False) As Point
        If lineA1.Z <> lineA2.Z OrElse lineA1.Z <> lineB1.Z OrElse lineA1.Z <> lineB2.Z Then
            Throw New ArgumentException("All points have to be in the same horizontal plane (z=constant)")
        End If

        If lineA1 = lineA2 Or lineB1 = lineB2 Then
            errorCode = ErrorCode.NoLineDetected
            Return Nothing
        ElseIf lineA1 = lineB2 And lineA2 = lineB1 Then
            errorCode = ErrorCode.UndefinedError
            Return Nothing
        ElseIf New Vector(lineA1, lineA2).ScaleOneVector = New Vector(lineB1, lineB2).ScaleOneVector Or New Vector(lineA1, lineA2).ScaleOneVector = New Vector(lineB1, lineB2).ScaleOneVector.Reverse Then
            errorCode = ErrorCode.ParallelLines
            If Distance(lineA1, lineB1, lineB2, infiniteLines) > 0 Then
                Return Nothing
            ElseIf Distance(lineA1, lineB1) / Distance(lineA1, lineA2) < 1 AndAlso Distance(lineA2, lineB1) / Distance(lineA1, lineA2) < 1 Then
                'OrElse Distance(lineA_1, lineB_2) / Distance(lineA_1, lineA_2) < 1 And Distance(lineA_2, lineB_2) / Distance(lineA_1, lineA_2) < 1
                If Distance(lineA1, lineB1) < Distance(lineA2, lineB1) Then
                    Return New Line(lineA1, lineB1).MidPoint
                Else
                    Return New Line(lineA2, lineB1).MidPoint
                End If
            ElseIf Distance(lineA1, lineB2) / Distance(lineA1, lineA2) < 1 And Distance(lineA2, lineB2) / Distance(lineA1, lineA2) < 1 Then
                If Distance(lineA1, lineB2) < Distance(lineA2, lineB2) Then
                    Return New Line(lineA1, lineB2).MidPoint
                Else
                    Return New Line(lineA2, lineB2).MidPoint
                End If
            ElseIf infiniteLines Then
                If Distance(lineA1, lineB1) < Distance(lineA2, lineB1) Then
                    If Distance(lineA1, lineB1) < Distance(lineA1, lineB2) Then
                        Return New Line(lineA1, lineB1).MidPoint
                    Else
                        Return New Line(lineA1, lineB2).MidPoint
                    End If
                Else
                    If Distance(lineA2, lineB1) < Distance(lineA2, lineB2) Then
                        Return New Line(lineA2, lineB1).MidPoint
                    Else
                        Return New Line(lineA2, lineB2).MidPoint
                    End If
                End If
            Else
                Return Nothing
            End If

#Region "PreviousCode"
            'ElseIf CheckIfHorizontal(lineA_1, lineA_2) And CheckIfVertical(lineB_1, lineB_2) Then
            '    errorCode = ErrorCode.noError
            '    Return New Point(lineB_1.X, lineA_1.Y)
            'ElseIf CheckIfVertical(lineA_1, lineA_2) And CheckIfHorizontal(lineB_1, lineB_2) Then
            '    errorCode = ErrorCode.noError
            '    Return New Point(lineA_1.X, lineB_1.Y)
#End Region

        Else
            Dim interPt As Point = Nothing
            'If infiniteLines Then
            If CheckIfVertical(lineA1, lineA2) Then
                Dim lineBC = (lineB2.Y - lineB1.Y) / (lineB2.X - lineB1.X)
                Dim lineBD = lineB1.Y - lineBC * lineB1.X
                'le point de retour est celui qui vérifie l'équation de la ligne B avec X = XA
                Dim xpoint = lineA1.X
                Dim ypoint = lineBC * xpoint + lineBD
                errorCode = ErrorCode.NoError
                interPt = New Point(xpoint, ypoint, lineA1.Z)

            ElseIf CheckIfVertical(lineB1, lineB2) Then
                Dim lineAC = (lineA2.Y - lineA1.Y) / (lineA2.X - lineA1.X)
                Dim lineAD = lineA1.Y - lineAC * lineA1.X
                'le point de retour est celui qui vérifie l'équation de la ligne B avec X = XB
                Dim xpoint = lineB1.X
                Dim ypoint = lineAC * xpoint + lineAD
                errorCode = ErrorCode.NoError
                interPt = New Point(xpoint, ypoint, lineA1.Z)

            Else
                'param équations des droites y=Cx+D
                Dim lineAC = (lineA2.Y - lineA1.Y) / (lineA2.X - lineA1.X)
                Dim lineAD = lineA1.Y - lineAC * lineA1.X
                Dim lineBC = (lineB2.Y - lineB1.Y) / (lineB2.X - lineB1.X)
                Dim lineBD = lineB1.Y - lineBC * lineB1.X
                'le point de retour est celui qui vérifie les 2 équations
                If lineAC = lineBC Then
                    errorCode = ErrorCode.ParallelLines
                    If Distance(lineA1, lineB1, lineB2, infiniteLines) > 0 OrElse Distance(lineA1, lineB1) / Distance(lineA1, lineA2) < 1 AndAlso Distance(lineA2, lineB1) / Distance(lineA1, lineA2) < 1 OrElse Distance(lineA1, lineB2) / Distance(lineA1, lineA2) < 1 And Distance(lineA2, lineB2) / Distance(lineA1, lineA2) < 1 Then
                        Return Nothing
                    Else
                        If Distance(lineA1, lineB1) < Distance(lineA2, lineB1) Then
                            If Distance(lineA1, lineB1) < Distance(lineA1, lineB2) Then
                                Return New Line(lineA1, lineB1).MidPoint
                            Else
                                Return New Line(lineA1, lineB2).MidPoint
                            End If
                        Else
                            If Distance(lineA2, lineB1) < Distance(lineA2, lineB2) Then
                                Return New Line(lineA2, lineB1).MidPoint
                            Else
                                Return New Line(lineA2, lineB2).MidPoint
                            End If
                        End If
                    End If
                End If
                Dim xpoint = (lineBD - lineAD) / (lineAC - lineBC)
                Dim ypoint = lineAC * xpoint + lineAD
                errorCode = ErrorCode.NoError
                interPt = New Point(xpoint, ypoint, lineA1.Z)
            End If

            If infiniteLines Then
                Return interPt
            Else
                If Distance(interPt, lineA1, lineA2, False) = 0.0 AndAlso Distance(interPt, lineB1, lineB2, False) = 0.0 Then
                    Return interPt
                Else
                    Return Nothing
                End If
            End If
#Region "PreviousCode"
            'Else
            '    If lineA_1 = lineB_1 Then
            '        Return lineA_1
            '    ElseIf lineA_1 = lineB_2 Then
            '        Return lineA_1
            '    ElseIf lineA_2 = lineB_1 Then
            '        Return lineA_2
            '    ElseIf lineA_2 = lineB_2 Then
            '        Return lineA_2
            '    Else
            '        Dim a As Double = (lineA_2.Y - lineA_1.Y) / (lineA_2.X - lineA_1.X)
            '        Dim b As Double = lineA_1.Y - a * CDbl(lineA_1.X)

            '        Dim c As Double = (lineB_2.Y - lineB_1.Y) / (lineB_2.X - lineB_1.X)
            '        Dim d As Double = lineB_1.Y - c * CDbl(lineB_1.X)

            '        If Not Double.IsInfinity(a) And Not Double.IsInfinity(c) And Not a = 0 And Not c = 0 Then
            '            errorCode = ErrorCode.noError
            '            Dim x As Double = (d - b) / (a - c)
            '            Return New Point(x, a * x + b)
            '        ElseIf Double.IsInfinity(a) Then
            '            errorCode = ErrorCode.noError
            '            Return New Point(CDbl(lineA_1.X), CDbl(lineA_1.X) * c + d)
            '        ElseIf Double.IsInfinity(c) Then
            '            errorCode = ErrorCode.noError
            '            Return New Point(CDbl(lineB_1.X), CDbl(lineB_1.X) * a + b)
            '        ElseIf a = 0 Then
            '            errorCode = ErrorCode.noError
            '            Return New Point((b - d) / c, b)
            '        ElseIf c = 0 Then
            '            errorCode = ErrorCode.noError
            '            Return New Point((d - b) / a, d)
            '        Else
            '            errorCode = ErrorCode.undefinedError
            '            Return Nothing
            '        End If
            'End If
            'End If
#End Region
        End If
    End Function

    ''' <summary>
    ''' Get the Intersection 3D Point of a Plane and a segment or an infinite line
    ''' </summary>
    ''' <param name="plan"> Plane </param>
    ''' <param name="pt1"> First Point of Line </param>
    ''' <param name="pt2"> Second Point of Line </param>
    ''' <param name="ec"> Error </param>
    ''' <param name="infiniteLine"> Infinite line or segment</param>
    ''' <returns></returns>
    Public Function GetIntersection(plan As Plane, pt1 As Point, pt2 As Point, ByRef ec As ErrorCode, Optional infiniteLine As Boolean = False) As Point
        If pt1 = pt2 Then
            ec = ErrorCode.NoLineDetected
            Return Nothing
        End If
        If plan.Vector.Length = 0 OrElse IsNothing(plan.Origin) Then
            ec = ErrorCode.NoPlanDetected
            Return Nothing
        End If
        If New Vector(pt1, pt2) * plan.Vector = 0 Then
            ec = ErrorCode.ParallelToPlan
            Return Nothing
        End If
        Dim scal1 = New Vector(plan.Origin, pt1) * plan.Vector
        Dim scal2 = New Vector(plan.Origin, pt2) * plan.Vector
        If scal1 = 0 Then
            ec = ErrorCode.NoError
            Return pt1
        End If
        If scal2 = 0 Then
            ec = ErrorCode.NoError
            Return pt2
        End If
        If scal1 * scal2 > 0 AndAlso Not infiniteLine Then
            ec = ErrorCode.NoError
            Return Nothing
        Else
            Dim k = ((plan.Origin.X - pt1.X) * plan.Vector.X + (plan.Origin.Y - pt1.Y) * plan.Vector.Y + (plan.Origin.Z - pt1.Z) * plan.Vector.Z) / ((pt2.X - pt1.X) * plan.Vector.X + (pt2.Y - pt1.Y) * plan.Vector.Y + (pt2.Z - pt1.Z) * plan.Vector.Z)
            ec = ErrorCode.NoError
            Return New Point() With {
            .X = pt1.X + k * (pt2.X - pt1.X),
            .Y = pt1.Y + k * (pt2.Y - pt1.Y),
            .Z = pt1.Z + k * (pt2.Z - pt1.Z)
            }
        End If
    End Function

    ''' <summary>
    ''' Check if a vector/segment/infinite line is horizontal in plan 2D(X|Y)
    ''' </summary>
    ''' <param name="pointA"></param>
    ''' <param name="pointB"></param>
    ''' <returns></returns>
    Private Function CheckIfHorizontal(pointA As Point, pointB As Point) As Boolean
        Return pointA.Y = pointB.Y
    End Function

    ''' <summary>
    ''' Check if a vector/segment/infinite line is vertical in plan 2D(X|Y)
    ''' </summary>
    ''' <param name="pointA"></param>
    ''' <param name="pointB"></param>
    ''' <returns></returns>
    Private Function CheckIfVertical(pointA As Point, pointB As Point) As Boolean
        Return pointA.X = pointB.X
    End Function

    'POur l'instant fonction non complète, la fusion est toujours faite (à quoi sert le fillHoleBetweenThem ?)
    ''' <summary>
    ''' Fusion of two Vectors
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="fillHoleBetweenThem"></param>
    ''' <returns></returns>
    Public Function MergeVectors(v1 As Vector, v2 As Vector, Optional fillHoleBetweenThem As Boolean = False) As Vector
        Dim canBeMerged As Boolean = fillHoleBetweenThem
        If Not canBeMerged Then

        End If

        Dim vectors As New List(Of Vector)
        Dim result As Vector = Nothing

        ' Add all possible combinations
        vectors.Add(New Vector(v1.Origin, v2.Origin))
        vectors.Add(New Vector(v1.Origin, v2.EndPoint))
        vectors.Add(New Vector(v1.EndPoint, v2.Origin))
        vectors.Add(New Vector(v1.EndPoint, v2.EndPoint))

        For Each vector In vectors
            If result Is Nothing Then
                result = vector
            ElseIf vector.Length > result.Length Then
                result = vector
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Projects 2D two segments on each other
    ''' </summary>
    ''' <param name="l1"> First segment </param>
    ''' <param name="l2"> Second segment </param>
    ''' <returns> Couple of two new segments </returns>
    Public Function GetCommonProjectedLines(l1 As Line, l2 As Line) As Tuple(Of Line, Line)
        '2d only
        Dim st1 As New Point(l1.StartPoint.X, l1.StartPoint.Y, 0)
        Dim end1 As New Point(l1.EndPoint.X, l1.EndPoint.Y, 0)
        Dim st2 As New Point(l2.StartPoint.X, l2.StartPoint.Y, 0)
        Dim end2 As New Point(l2.EndPoint.X, l2.EndPoint.Y, 0)

        Dim projectedPoint1 As Point = GetNormalPoint(st2, st1, end1)
        Dim projectedPoint2 As Point = GetNormalPoint(end2, st1, end1)
        Dim projectedLine1 As New Line(projectedPoint1, projectedPoint2)

        Dim pDist1 = Math.Round(Distance(projectedPoint1, st1, end1, False), -DefaultTolerance) 'Bricolage : arrondi rajouté pour corriger bug sur datcha lancée sur plusieurs bâtiments en même temps
        Dim pDist2 = Math.Round(Distance(projectedPoint2, st1, end1, False), -DefaultTolerance)

        If pDist1 > 0 And pDist2 > 0 Then
            If CheckIfPointBelongsToLine(st1, projectedLine1, Nothing) And CheckIfPointBelongsToLine(end1, projectedLine1, Nothing) Then
                projectedLine1 = New Line(st1, end1)
            Else
                projectedLine1 = Nothing
            End If
        ElseIf pDist1 = 0 And pDist2 > 0 Then
            If Distance(projectedPoint2, st1) < Distance(projectedPoint2, end1) Then
                projectedLine1 = New Line(st1, projectedPoint1)
            Else
                projectedLine1 = New Line(projectedPoint1, end1)
            End If
        ElseIf pDist1 > 0 And pDist2 = 0 Then
            If Distance(projectedPoint1, st1) < Distance(projectedPoint1, end1) Then
                projectedLine1 = New Line(st1, projectedPoint2)
            Else
                projectedLine1 = New Line(projectedPoint2, end1)
            End If
        End If
        If projectedLine1 IsNot Nothing AndAlso Distance(projectedLine1.EndPoint, st1) < Distance(projectedLine1.StartPoint, st1) Then
            projectedLine1 = projectedLine1.Reverse
        End If

        Dim projectedPoint3 As Point = GetNormalPoint(st1, st2, end2)
        Dim projectedPoint4 As Point = GetNormalPoint(end1, st2, end2)
        Dim projectedLine2 As New Line(projectedPoint3, projectedPoint4)

        Dim pDist3 = Distance(projectedPoint3, st2, end2, False)
        Dim pDist4 = Distance(projectedPoint4, st2, end2, False)

        If pDist3 > 0 And pDist4 > 0 Then
            If CheckIfPointBelongsToLine(st2, projectedLine2, Nothing) And CheckIfPointBelongsToLine(end2, projectedLine2, Nothing) Then
                projectedLine2 = New Line(st2, end2)
            Else
                projectedLine2 = Nothing
            End If
        ElseIf pDist3 = 0 And pDist4 > 0 Then
            If Distance(projectedPoint4, st2) < Distance(projectedPoint4, end2) Then
                projectedLine2 = New Line(st2, projectedPoint3)
            Else
                projectedLine2 = New Line(projectedPoint3, end2)
            End If
        ElseIf pDist3 > 0 And pDist4 = 0 Then
            If Distance(projectedPoint3, st2) < Distance(projectedPoint3, end2) Then
                projectedLine2 = New Line(st2, projectedPoint4)
            Else
                projectedLine2 = New Line(projectedPoint4, end2)
            End If
        End If
        If projectedLine2 IsNot Nothing AndAlso Distance(projectedLine2.EndPoint, st2) < Distance(projectedLine2.StartPoint, st2) Then
            projectedLine2 = projectedLine2.Reverse
        End If

        Return New Tuple(Of Line, Line)(projectedLine1, projectedLine2)
    End Function

    ''' <summary>
    ''' Check 3D if point P belongs to segment L
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="l"> Line </param>
    ''' <param name="tolerance"> Set the Tolerance as 10^-tolerance (evaluate the vectors proper coordinates), -DefaultTolerance is used by default.. </param>
    ''' <returns></returns>
    <Obsolete("Use CheckIfPointBelongsToLine")>
    Public Function CheckIfPointIsInLine(p As Point, l As Line, Optional tolerance As Integer? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = -DefaultTolerance
        Dim checkVector As New Vector(l.StartPoint, p)
        If Vector.AreAligned(checkVector, l.Vector, -tolerance) Then
            Dim k = Distance(l.StartPoint, p) / l.Length
            Dim q = Distance(l.EndPoint, p) / l.Length
            If k >= 0 And k <= 1 And q >= 0 And q <= 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Check 3D if point P belongs to segment/infiniteLine L
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="l"> Line </param>
    ''' <param name="tolerance"> Set the Tolerance as 10^tolerance (evaluate the vectors proper coordinates), DefaultTolerance is used by default.. </param>
    ''' <returns></returns>
    Public Function CheckIfPointBelongsToLine(p As Point, l As Line, Optional tolerance As Integer? = Nothing, Optional infiniteLine As Boolean = False) As Boolean
        If tolerance Is Nothing Then tolerance = DefaultTolerance
        If infiniteLine Then
            Return Distance(p, l, True) <= Math.Pow(10, tolerance)
        End If
        Return Distance(p, l.Vector, False) <= Math.Pow(10, tolerance)
    End Function

    ''' <summary>
    ''' Check 3D if point P belongs to segment/infiniteLine L
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="l"> Line </param>
    ''' <param name="tolerance"> Set the Tolerance(evaluate the vectors proper coordinates), 10^DefaultTolerance is used by default.. </param>
    ''' <returns></returns>
    Public Function CheckIfPointBelongsToLine(p As Point, l As Line, Optional tolerance As Double? = Nothing, Optional infiniteLine As Boolean = False) As Boolean
        If tolerance Is Nothing Then tolerance = Math.Pow(10, DefaultTolerance)
        If infiniteLine Then
            Return Distance(p, l.Vector, True) <= tolerance
        End If
        Return Distance(p, l.Vector, False) <= tolerance
    End Function

    ''' <summary>
    ''' Check 3D if distance from point P to segment/infiniteLine L is less than D without P belonging to L
    ''' </summary>
    ''' <param name="p"> Point </param>
    ''' <param name="l"> Line </param>
    ''' <param name="d"> Distance </param>
    ''' <returns></returns>
    Public Function CheckIfPointIsNearLine(p As Point, l As Line, d As Double, Optional infiniteLine As Boolean = False) As Boolean
        Dim checkVector1 As New Vector(l.StartPoint, p)
        Dim checkVector2 As New Vector(l.EndPoint, p)
        If Distance(p, l.Vector, False) <= d Then
            Dim tempTol = DefaultTolerance
            While Math.Pow(10, tempTol + 1) >= d
                tempTol -= 1
            End While
            Return Not CheckIfPointBelongsToLine(p, l, tempTol, infiniteLine)
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Matrix transformation of a Point
    ''' </summary>
    ''' <param name="point"> Point to transform </param>
    ''' <param name="m"> Matrix for transformation </param>
    ''' <returns> New Point </returns>
    Public Function TransformMat4(point As Point, m As Double(,)) As Point
        Dim transformedPoint As New Point()
        Dim x = point.X, y = point.Y, z = point.Z
        Dim w = m(0, 3) * x + m(1, 3) * y + m(2, 3) * z + m(3, 3)
        transformedPoint.X = (m(0, 0) * x + m(1, 0) * y + m(2, 0) * z + m(3, 0)) / w
        transformedPoint.Y = (m(0, 1) * x + m(1, 1) * y + m(2, 1) * z + m(3, 1)) / w
        transformedPoint.Z = (m(0, 2) * x + m(1, 2) * y + m(2, 2) * z + m(3, 2)) / w
        Return transformedPoint
    End Function

    ''' <summary>
    ''' Matrix 3x3 multiplication
    ''' </summary>
    ''' <param name="m1"></param>
    ''' <param name="m2"></param>
    ''' <returns></returns>
    Public Function MultiplyMat3(m1 As Double(,), m2 As Double(,)) As Double(,)
        Dim out(2, 2) As Double
        'first column
        out(0, 0) = RoundToSignificantDigits(m1(0, 0) * m2(0, 0) + m1(1, 0) * m2(0, 1) + m1(2, 0) * m2(0, 2), RealPrecision)
        out(0, 1) = RoundToSignificantDigits(m1(0, 1) * m2(0, 0) + m1(1, 1) * m2(0, 1) + m1(2, 1) * m2(0, 2), RealPrecision)
        out(0, 2) = RoundToSignificantDigits(m1(0, 2) * m2(0, 0) + m1(1, 2) * m2(0, 1) + m1(2, 2) * m2(0, 2), RealPrecision)
        'second column
        out(1, 0) = RoundToSignificantDigits(m1(0, 0) * m2(1, 0) + m1(1, 0) * m2(1, 1) + m1(2, 0) * m2(1, 2), RealPrecision)
        out(1, 1) = RoundToSignificantDigits(m1(0, 1) * m2(1, 0) + m1(1, 1) * m2(1, 1) + m1(2, 1) * m2(1, 2), RealPrecision)
        out(1, 2) = RoundToSignificantDigits(m1(0, 2) * m2(1, 0) + m1(1, 2) * m2(1, 1) + m1(2, 2) * m2(1, 2), RealPrecision)
        'third column
        out(2, 0) = RoundToSignificantDigits(m1(0, 0) * m2(2, 0) + m1(1, 0) * m2(2, 1) + m1(2, 0) * m2(2, 2), RealPrecision)
        out(2, 1) = RoundToSignificantDigits(m1(0, 1) * m2(2, 0) + m1(1, 1) * m2(2, 1) + m1(2, 1) * m2(2, 2), RealPrecision)
        out(2, 2) = RoundToSignificantDigits(m1(0, 2) * m2(2, 0) + m1(1, 2) * m2(2, 1) + m1(2, 2) * m2(2, 2), RealPrecision)
        Return out
    End Function

    ''' <summary>
    ''' Check 2D if two Surfaces intersect each other
    ''' </summary>
    ''' <param name="surf1"> First Surface </param>
    ''' <param name="surf2"> Second Surface </param>
    ''' <returns></returns>
    Public Function DoSurfacesIntersect(surf1 As Surface, surf2 As Surface) As Boolean 'Ajouter une tolerance
        If surf1.Type = Surface.SurfaceType.Circle Xor surf1.Type = Surface.SurfaceType.Circle Then
            If surf1.Type = Surface.SurfaceType.Circle Then
                If Distance(surf1.Points(0), surf2) <= Distance(surf1.Points(0), surf1.Points(1)) Then
                    Return True
                Else
                    Return False
                End If
            Else
                If Distance(surf2.Points(0), surf1) <= Distance(surf2.Points(0), surf2.Points(1)) Then
                    Return True
                Else
                    Return False
                End If
            End If
        ElseIf surf1.Type = Surface.SurfaceType.Circle Then
            If Distance(surf1.Points(0), surf2.Points(0)) <= Distance(surf1.Points(0), surf1.Points(1)) + Distance(surf2.Points(0), surf2.Points(1)) Then
                Return True
            Else
                Return False
            End If
        Else
            For Each point In surf1.Points
                If surf2.WindingNumber(point) Mod 2 = 1 Then
                    Return True
                End If
            Next
            For Each point In surf2.Points
                If surf1.WindingNumber(point) Mod 2 = 1 Then
                    Return True
                End If
            Next
            Return False
        End If
    End Function

#Region "ToCheck"
    'Good
    ''' <summary>
    ''' Express point's coordinates in a new axis system (must be orthonormal,at least  orthogonal).
    ''' </summary>
    ''' <param name="pointCoord"> Point's coordinates to express in new axis system (actual system axis). </param>
    ''' <param name="axisOrigin"> New origin Point's coordinates (in actual axis system).</param>
    ''' <param name="axisVectX"> New fisrt axis's coordinates (in actual axis system).</param>
    ''' <param name="axisVectY"> New second axis's coordinates (in actual axis system).</param>
    ''' <param name="axisVectZ"> New third axis's coordinates (in actual axis system).</param>
    ''' <returns> New point's coordinates </returns>
    Public Function ChangeCoordPointInAnotherAxisSyst(ByVal pointCoord(), ByVal axisOrigin(), ByVal axisVectX(), ByVal axisVectY(), ByVal axisVectZ())
        'these vectors need to be normalized
        NormerVecteur(axisVectX, axisVectX)
        NormerVecteur(axisVectY, axisVectY)
        NormerVecteur(axisVectZ, axisVectZ)

        'loop through all the points getting each coordinate, and transform it into the axis system coordinate
        Dim delta(2)
        Dim csCoords(2)

        'vector from cs origin to point
        delta(0) = pointCoord(0) - axisOrigin(0)
        delta(1) = pointCoord(1) - axisOrigin(1)
        delta(2) = pointCoord(2) - axisOrigin(2)
        'determine components along cs axis
        csCoords(0) = RoundToSignificantDigits(DotProduct(delta, axisVectX), RealPrecision)
        csCoords(1) = RoundToSignificantDigits(DotProduct(delta, axisVectY), RealPrecision)
        csCoords(2) = RoundToSignificantDigits(DotProduct(delta, axisVectZ), RealPrecision)
        Return csCoords
    End Function

    ''' <summary>
    ''' Express point's coordinates in a new axis system (must be orthonormal,at least  orthogonal).
    ''' </summary>
    ''' <param name="point"> Point to express in new axis system (actual system axis). </param>
    ''' <param name="axisOrigin"> New origin Point (in actual axis system).</param>
    ''' <param name="axisVectX"> New fisrt axis (in actual axis system).</param>
    ''' <param name="axisVectY"> New second axis (in actual axis system).</param>
    ''' <param name="axisVectZ"> New third axis (in actual axis system).</param>
    ''' <returns> New point's coordinates </returns>
    Public Function ChangeCoordPointInAnotherAxisSyst(point As Point, axisOrigin As Point, axisVectX As Vector, axisVectY As Vector, axisVectZ As Vector) As Point
        Dim newCoords = ChangeCoordPointInAnotherAxisSyst({point.X, point.Y, point.Z}, {axisOrigin.X, axisOrigin.Y, axisOrigin.Z},
                                                          {axisVectX.X, axisVectX.Y, axisVectX.Z}, {axisVectY.X, axisVectY.Y, axisVectY.Z},
                                                          {axisVectZ.X, axisVectZ.Y, axisVectZ.Z})
        Return New Point(newCoords(0), newCoords(1), newCoords(2))
    End Function

    'Good comme le ScalarProduct mais fonctionnne également entre un point et un vecteur
    ''' <summary>
    ''' Scalar product
    ''' </summary>
    ''' <param name="vect1"> First Vector/Point </param>
    ''' <param name="vect2"> Second Vector/Point </param>
    ''' <returns></returns>
    Public Function DotProduct(vect1(), vect2()) As Double
        DotProduct = RoundToSignificantDigits(vect1(0) * vect2(0) + vect1(1) * vect2(1) + vect1(2) * vect2(2), RealPrecision)
    End Function

    'Good
    'Même chose que le scaleOneVector
    Public Sub NormerVecteur(invect(), ByRef normvect())
        If invect(0) = 0 And invect(1) = 0 And invect(2) = 0 Then
            Debug.Print("Stop")
        End If
        Dim mag 'As Double
        mag = Math.Sqrt(invect(0) * invect(0) + invect(1) * invect(1) + invect(2) * invect(2))
        If mag = 0 Then
            Call Err.Raise(1001, , "Zero length vector cannot be normalized")
        End If
        normvect(0) = RoundToSignificantDigits(invect(0) / mag, RealPrecision)
        normvect(1) = RoundToSignificantDigits(invect(1) / mag, RealPrecision)
        normvect(2) = RoundToSignificantDigits(invect(2) / mag, RealPrecision)
    End Sub
#End Region

#Region "DATA Structure"
    'Intersection, Union,Différence de polygones

    ''' <summary>
    ''' Data Structures for two polygons to execute clipping/intersection/fusion of them by Greiner/Hormann method
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="c"></param>
    ''' <returns></returns>
    Public Function CreateDataStructure(s As Surface, c As Surface) As Tuple(Of List(Of VertexData), List(Of VertexData))
        Dim clist, slist As New List(Of VertexData)

        If s.Points.Count = 0 OrElse c.Points.Count = 0 Then
            Throw New Exception("Empty surface(s)")
            GoTo Vide
        End If

        'Step 0 : Retournement du polygone C si nécessaire, le subject et le clipping ne doivent pas tourner dans le même sens
        If Not (s.GetInnerAngles().Item2) Xor c.GetInnerAngles().Item2 Then
            Dim cprim As New List(Of Point)
            For i = 1 To c.Points.Count
                cprim.Add(c.Points(c.Points.Count - i))
            Next
            c = New Surface(cprim, Surface.SurfaceType.Polygon)
        End If

        'Step 1 : Création des liste de vertex des polygones et de leurs intersections communes
        For i = 0 To s.Points.Count - 1
            Dim s1 = New VertexData(s.Points(i))
            If Not SelectVertexIfExistInList(slist, s1) Then
                slist.Add(s1)
            End If

            Dim s2 As VertexData
            If i <> (s.Points.Count - 1) Then
                s2 = New VertexData(s.Points(i + 1))
            Else
                s2 = New VertexData(s.Points(0))
            End If
            If Not (SelectVertexIfExistInList(slist, s2)) Then
                slist.Add(s2)
            End If

            s1.Suiv = s2
            s2.Prec = s1

            For j = 0 To c.Points.Count - 1
                Dim c1 = New VertexData(c.Points(j))
                If Not (SelectVertexIfExistInList(clist, c1)) Then
                    clist.Add(c1)
                End If

                Dim c2 As VertexData
                If j <> (c.Points.Count - 1) Then
                    c2 = New VertexData(c.Points(j + 1))
                Else
                    c2 = New VertexData(c.Points(0))
                End If
                If Not (SelectVertexIfExistInList(clist, c2)) Then
                    clist.Add(c2)
                End If

                If IsNothing(c1.Suiv) Then
                    c1.Suiv = c2
                End If
                If IsNothing(c2.Prec) Then
                    c2.Prec = c1
                End If

                Dim alphaS, alphaC As New Double
                Dim pointI = VerticesIntersection(s1.Point, s2.Point, c1.Point, c2.Point, alphaS, alphaC)
                If pointI IsNot Nothing Then
                    'Dim vectorS = New Vector(S1.point, S2.point)
                    'Dim vectorC = New Vector(C1.point, C2.point)

                    Dim i1 As VertexData
                    Dim i2 As VertexData
                    If alphaS = 0 Then
                        i1 = s1
                        i1.Intersect = True
                        i1.Alpha = alphaS
                    ElseIf alphaS = 1 Then
                        i1 = s2
                        i1.Intersect = True
                        i1.Alpha = alphaS
                    Else
                        'Dim pointI = vectorS.GetPointAtDistance(alphaS, S1.point)
                        i1 = New VertexData(pointI, s1, s2, True) With {
                            .Alpha = alphaS
                        }
                        VertexSortOnList(slist, i1)
                    End If

                    If alphaC = 0 Then
                        i2 = c1
                        i2.Intersect = True
                        i2.Alpha = alphaC
                    ElseIf alphaC = 1 Then
                        i2 = c2
                        i2.Intersect = True
                        i2.Alpha = alphaC
                    Else
                        i2 = New VertexData(i1.Point, c1, c2, True) With {
                            .Alpha = alphaC
                        }
                        VertexSortOnList(clist, i2)
                    End If
                    i1.Neighbour = i2
                    i2.Neighbour = i1

                End If
            Next
        Next

        'Step 2 : Attribution d'une valeur d'entrée ou sortie au vertex d'intersection
        Dim sDedouble As New List(Of VertexData)

        'Pour Slist
        For Each vertex In slist
            If vertex.Intersect Then
                Dim testPoint = New Line(vertex.Point, vertex.Suiv.Point).MidPoint
                Dim fakeVertex = New VertexData(testPoint)
                vertex.EntryExit = fakeVertex.IsVertexInsidePolygonVertices(clist)
                vertex.Neighbour.EntryExit = vertex.EntryExit

                Dim testPoint2 = New Line(vertex.Prec.Point, vertex.Point).MidPoint
                Dim fakeVertex2 = New VertexData(testPoint2)
                If fakeVertex2.IsVertexInsidePolygonVertices(clist) = vertex.EntryExit Then
                    sDedouble.Add(vertex)
                End If
            End If
        Next

        'If SDedouble.Count > 0 Then
        '    For Each vertexToDedouble In SDedouble
        '        Dim newVertex = New VertexData(New Point(vertexToDedouble.point.X + 1, vertexToDedouble.point.Y), vertexToDedouble.prec, vertexToDedouble, True) With {
        '        .alpha = vertexToDedouble.prec.alpha + 0.1,
        '        .entry_exit = Not vertexToDedouble.entry_exit
        '        }
        '        Dim neighbourToDedouble = vertexToDedouble.neighbour
        '        Dim newNeighbour = New VertexData(New Point(vertexToDedouble.point.X + 1, vertexToDedouble.point.Y), neighbourToDedouble.prec, neighbourToDedouble, True) With {
        '        .alpha = neighbourToDedouble.prec.alpha + 0.1,
        '        .entry_exit = Not neighbourToDedouble.entry_exit
        '        }
        '        VertexSortOnList(Slist, newVertex)
        '        VertexSortOnList(Clist, newNeighbour)

        '        newVertex.alpha = vertexToDedouble.alpha
        '        newVertex.point.X -= 1
        '        newNeighbour.alpha = neighbourToDedouble.alpha
        '        newNeighbour.point.X -= 1

        '        newVertex.neighbour = newNeighbour
        '        newNeighbour.neighbour = newVertex
        '    Next
        'End If
Vide:
        Return New Tuple(Of List(Of VertexData), List(Of VertexData))(slist, clist)
    End Function

    ''' <summary>
    ''' Check if vertex exists in list of vertices.
    ''' </summary>
    ''' <param name="vertices"></param>
    ''' <param name="vertex"></param>
    ''' <returns></returns>
    Public Function SelectVertexIfExistInList(vertices As List(Of VertexData), ByRef vertex As VertexData) As Boolean
        Dim result = False
        If Not IsNothing(vertices) Then
            For Each elem In vertices
                If elem.Point = vertex.Point Then
                    result = True
                    vertex = elem
                    Exit For
                End If
            Next
        End If
        Return result
    End Function

    ''' <summary>
    ''' Intersection bewteen two lines of vertices.
    ''' </summary>
    ''' <param name="p1"> First vertex of Line P. </param>
    ''' <param name="p2"> Second vertex of Line P. </param>
    ''' <param name="q1"> First vertex of Line Q. </param>
    ''' <param name="q2"> Second vertex of Line Q. </param>
    ''' <param name="alphaP"> Parameter of the intersect Vertex on Line P. </param>
    ''' <param name="alphaQ"> Parameter of the intersect Vertex on Line Q. </param>
    ''' <returns></returns>
    Public Function VerticesIntersection(p1 As Point, p2 As Point, q1 As Point, q2 As Point, ByRef alphaP As Double, ByRef alphaQ As Double) As Point
        Dim result = Nothing
        Dim ec = New ErrorCode
        Dim inter = GetIntersection(p1, p2, q1, q2, ec)
        If ec = ErrorCode.NoError AndAlso inter IsNot Nothing Then
            alphaP = Distance(p1, inter) / Distance(p1, p2)
            alphaQ = Distance(q1, inter) / Distance(q1, q2)
            result = inter
        End If
        'Dim WEC_P1 = (P1.X - Q1.X) * (Q2.Y - Q1.Y) - (P1.Y - Q1.Y) * (Q2.X - Q1.X)
        'Dim WEC_P2 = (P2.X - Q1.X) * (Q2.Y - Q1.Y) - (P2.Y - Q1.Y) * (Q2.X - Q1.X)
        'If WEC_P1 * WEC_P2 <= 0 AndAlso Not (WEC_P1 = 0 AndAlso WEC_P2 = 0) Then
        '    Dim WEC_Q1 = (Q1.X - P1.X) * (P2.Y - P1.Y) - (Q1.Y - P1.Y) * (P2.X - P1.X)
        '    Dim WEC_Q2 = (Q2.X - P1.X) * (P2.Y - P1.Y) - (Q2.Y - P1.Y) * (P2.X - P1.X)
        '    If WEC_Q1 * WEC_Q2 <= 0 AndAlso Not (WEC_Q1 = 0 AndAlso WEC_Q2 = 0) Then
        '        alphaP = RoundToSignificantDigits(WEC_P1 / (WEC_P1 - WEC_P2), RealPrecision)
        '        alphaQ = RoundToSignificantDigits(WEC_Q1 / (WEC_Q1 - WEC_Q2), RealPrecision)
        '        result = True
        '    End If
        'Else
        '    If P1 = Q1 AndAlso Distance(P2, Q1, Q2, False) > 0 Then
        '        alphaP = 0
        '        alphaQ = 0
        '        result = True
        '    ElseIf P1 = Q2 AndAlso Distance(P2, Q1, Q2, False) > 0 Then
        '        alphaP = 0
        '        alphaQ = 1
        '        result = True
        '    ElseIf P2 = Q1 AndAlso Distance(P1, Q1, Q2, False) > 0 Then
        '        alphaP = 1
        '        alphaQ = 0
        '        result = True
        '    ElseIf P2 = Q2 AndAlso Distance(P1, Q1, Q2, False) > 0 Then
        '        alphaP = 1
        '        alphaQ = 1
        '        result = True
        '    End If
        'End If
        Return (result)
    End Function

    ''' <summary>
    ''' Insert vertex in a list of vertices.
    ''' </summary>
    ''' <param name="vertices"></param>
    ''' <param name="vertex"></param>
    ''' <returns></returns>
    Public Function VertexSortOnList(ByRef vertices As List(Of VertexData), ByRef vertex As VertexData)
        Dim alreadyIn = False
        For Each existVertex In vertices
            If existVertex = vertex Then
                alreadyIn = True
                vertex = existVertex
                vertex.Intersect = True
                Exit For
            End If
        Next
        If Not alreadyIn Then
            Dim range = vertices.IndexOf(vertex.Suiv)
            Dim isIntersection = True
            While isIntersection
                Dim rangeprec As New Integer
                If range <> 0 Then
                    rangeprec = range - 1
                Else
                    rangeprec = vertices.Count - 1
                End If

                isIntersection = vertices(rangeprec).Intersect
                If isIntersection And vertices(rangeprec).Alpha > vertex.Alpha Then
                    range = rangeprec
                Else
                    vertex.Prec = vertices(rangeprec)
                    vertices(rangeprec).Suiv = vertex
                    vertex.Suiv = vertices(range)
                    vertices(range).Prec = vertex
                    If range <> 0 Then
                        vertices.Insert(range, vertex)
                    Else
                        vertices.Add(vertex)
                    End If
                    isIntersection = False
                End If
            End While
        End If
        Return 0
    End Function

    ''' <summary>
    ''' Intersection of Two surfaces.
    ''' </summary>
    ''' <param name="poly1"></param>
    ''' <param name="poly2"></param>
    ''' <returns> A list of surfaces </returns>
    Public Function SurfaceIntersection(poly1 As Surface, poly2 As Surface) As List(Of Surface)
        Dim result As New List(Of Surface)
        If poly1.Points.Count = 0 OrElse poly2.Points.Count = 0 Then
            Throw New Exception("Empty surface(s)")
            GoTo Vide
        End If

        If poly1.Type = Surface.SurfaceType.Polygon AndAlso poly2.Type = Surface.SurfaceType.Polygon Then
            If poly1 = poly2 Then
                result.Add(poly1)
                Return result
            End If
            Dim t = CreateDataStructure(poly1, poly2)
            Dim data1 = t.Item1
            Dim data2 = t.Item2

            Dim nbPtsInter = 0
            For Each vertex In data1
                If vertex.Intersect Then
                    nbPtsInter += 1
                End If
            Next

            If nbPtsInter = 0 Then
                'Vérification d'imbrication
                If data1(0).IsVertexInsidePolygonVertices(data2) Then
                    result.Add(poly1)
                ElseIf data2(0).IsVertexInsidePolygonVertices(data1) Then
                    result.Add(poly2)
                End If
            Else 'If nbPtsInter Mod 2 = 0 Then
                Dim controlList As New List(Of VertexData)
                Dim index = 0
                'Dim max = Data1.Count
                'While index < max
                While index < data1.Count AndAlso controlList.Count < data1.Count
                    If data1(index).Intersect AndAlso Not controlList.Contains(data1(index)) Then
                        Dim currentList As New List(Of Point)
                        Dim first = data1(index)
                        currentList.Add(first.Point)
                        Dim current = first
                        If Not controlList.Contains(current) Then
                            controlList.Add(current)
                        End If
                        Dim isData1 = True
                        Do
                            If current.EntryExit Then
                                Do
                                    current = current.Suiv
                                    If current <> first Then
                                        currentList.Add(current.Point)
                                        If isData1 AndAlso Not controlList.Contains(current) Then
                                            controlList.Add(current)
                                        End If
                                    End If
                                Loop Until current.Intersect AndAlso Not current.EntryExit
                                'If current <> first Then
                                '    index = Data1.IndexOf(current)
                                'End If
                            Else
                                Do
                                    current = current.Prec
                                    If current <> first Then
                                        currentList.Add(current.Point)
                                        If isData1 AndAlso Not controlList.Contains(current) Then
                                            controlList.Add(current)
                                        End If
                                    End If
                                Loop Until current.Intersect AndAlso current.EntryExit
                                'If current <> first Then
                                '    max = Data1.IndexOf(current)
                                'End If
                            End If
                            current = current.Neighbour
                            isData1 = Not isData1
                        Loop Until current = first
                        If currentList.Count >= 3 Then
                            result.Add(New Surface(currentList, Surface.SurfaceType.Polygon).CleanPoints())
                        End If
                    Else
                        If Not controlList.Contains(data1(index)) Then
                            controlList.Add(data1(index))
                        End If
                    End If
                    index += 1
                End While
            End If
        Else
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If
Vide:
        Return result
    End Function

    Public Function SurfacePolyBoolIntersection(poly1 As Surface, poly2 As Surface) As List(Of Surface)
        Dim result As New List(Of Surface)
        If poly1.Points.Count = 0 OrElse poly2.Points.Count = 0 Then
            Throw New Exception("Empty surface(s)")
            GoTo Vide
        End If
        If Not poly1.Type = Surface.SurfaceType.Polygon OrElse Not poly2.Type = Surface.SurfaceType.Polygon Then
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If

        Dim polyOp = New PolyBool

        Dim polyL1 As New PointList From {}
        Dim polyL2 As New PointList From {}

        For Each pt In poly1.Points
            Dim ppt = New PolyBoolCS.Point(pt.X, pt.Y)
            polyL1.Add(ppt)
        Next
        For Each pt In poly2.Points
            Dim ppt = New PolyBoolCS.Point(pt.X, pt.Y)
            polyL2.Add(ppt)
        Next
        Dim polygon1 = New Polygon

        polygon1.regions = New List(Of PointList)() From {polyL1}
        Dim polygon2 = New Polygon
        polygon2.regions = New List(Of PointList)() From {polyL2}

        Dim polygonIntersect = polyOp.intersect(polygon1, polygon2)

        For Each inter In polygonIntersect.regions
            Dim currList As New List(Of Point) From {}
            For Each ppt In inter
                currList.Add(New Point(ppt.x, ppt.y))
            Next
            result.Add(New Surface(currList, Surface.SurfaceType.Polygon))
        Next
Vide:
        Return result
    End Function

    ''' <summary>
    ''' Union of Two surfaces.
    ''' </summary>
    ''' <param name="poly1"></param>
    ''' <param name="poly2"></param>
    ''' <returns> A list of surfaces </returns>
    Public Function SurfaceUnion(poly1 As Surface, poly2 As Surface) As List(Of Surface)
        Dim result As New List(Of Surface)
        If poly1.Type = Surface.SurfaceType.Polygon AndAlso poly2.Type = Surface.SurfaceType.Polygon Then
            If poly1 = poly2 Then
                result.Add(poly1)
                Return result
            End If
            Dim t = CreateDataStructure(poly1, poly2)
            Dim data1 = t.Item1
            Dim data2 = t.Item2

            Dim nbPtsInter = 0
            For Each vertex In data1
                If vertex.Intersect Then
                    nbPtsInter += 1
                End If
            Next

            If nbPtsInter = 0 Then
                'Vérification d'imbrication
                If data1(0).IsVertexInsidePolygonVertices(data2) Then
                    result.Add(poly2)
                ElseIf data2(0).IsVertexInsidePolygonVertices(data1) Then
                    result.Add(poly1)
                Else
                    result.Add(poly1)
                    result.Add(poly2)
                End If
            ElseIf nbPtsInter Mod 2 = 0 Then
                Dim controlList = New List(Of VertexData)
                Dim index = 0
                'While index < Data1.Count
                While index < data1.Count AndAlso controlList.Count < data1.Count
                    If (data1(index).Intersect OrElse Not data1(index).IsVertexInsidePolygonVertices(data2)) AndAlso Not controlList.Contains(data1(index)) Then
                        Dim currentList As New List(Of Point)
                        Dim first = data1(index)
                        currentList.Add(first.Point)
                        Dim current = first
                        If Not controlList.Contains(current) Then
                            controlList.Add(current)
                        End If
                        'Dim isData1 = True
                        Do
                            If current.Intersect AndAlso current.EntryExit Then
                                current = current.Neighbour
                                Do
                                    current = current.Prec
                                    If current <> first Then
                                        currentList.Add(current.Point)
                                    End If
                                Loop Until current.Intersect
                                current = current.Neighbour
                                If current <> first Then
                                    index = data1.IndexOf(current)
                                End If
                            Else
                                current = current.Suiv
                                If current <> first Then
                                    currentList.Add(current.Point)
                                    If Not controlList.Contains(current) Then
                                        controlList.Add(current)
                                    End If
                                End If
                                index += 1
                            End If
                        Loop Until current = first
                        If currentList.Count >= 3 Then
                            result.Add(New Surface(currentList, Surface.SurfaceType.Polygon).CleanPoints())
                        End If
                    Else
                        If Not controlList.Contains(data1(index)) Then
                            controlList.Add(data1(index))
                        End If
                    End If
                    index += 1
                End While
            End If
        Else
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If
        Return result
    End Function

    Public Function SurfacePolyBoolUnion(poly1 As Surface, poly2 As Surface, Optional tolerance As Integer? = Nothing) As List(Of Surface)
        Dim tolerance2 As Integer
        If tolerance Is Nothing Then
            tolerance2 = DefaultTolerance
        Else
            tolerance2 = tolerance
        End If

        Dim result As New List(Of Surface)
        If poly1.Points.Count = 0 OrElse poly2.Points.Count = 0 Then
            Throw New Exception("Empty surface(s)")
            GoTo Vide
        End If
        If Not poly1.Type = Surface.SurfaceType.Polygon OrElse Not poly2.Type = Surface.SurfaceType.Polygon Then
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If

        Dim polyOp = New PolyBool

        Dim polyL1 As New PointList From {}
        Dim polyL2 As New PointList From {}

        For Each pt In poly1.Points
            Dim ppt = New PolyBoolCS.Point(Math.Round(pt.X, -tolerance2), Math.Round(pt.Y, -tolerance2))
            polyL1.Add(ppt)
        Next
        For Each pt In poly2.Points
            Dim ppt = New PolyBoolCS.Point(Math.Round(pt.X, -tolerance2), Math.Round(pt.Y, -tolerance2))
            polyL2.Add(ppt)
        Next
        Dim polygon1 = New Polygon
        polygon1.regions = New List(Of PointList)() From {polyL1}
        Dim polygon2 = New Polygon
        polygon2.regions = New List(Of PointList)() From {polyL2}

        Dim polygonDifference = polyOp.union(polygon1, polygon2)

        For Each inter In polygonDifference.regions
            Dim currList As New List(Of Point) From {}
            For Each ppt In inter
                currList.Add(New Point(ppt.x, ppt.y, poly1.Points(0).Z)) 'SET POLY1 ELEVATION TO RESULT
            Next
            result.Add(New Surface(currList, Surface.SurfaceType.Polygon))
        Next
Vide:
        Return result
    End Function

    ''' <summary>
    ''' Difference between two surfaces.
    ''' </summary>
    ''' <param name="poly1"></param>
    ''' <param name="poly2"></param>
    ''' <returns></returns>
    Public Function SurfaceDifference(poly1 As Surface, poly2 As Surface) As List(Of Surface)
        Dim result As New List(Of Surface)
        If poly1.Type = Surface.SurfaceType.Polygon AndAlso poly2.Type = Surface.SurfaceType.Polygon Then
            If poly1 = poly2 Then
                Return Nothing
            End If
            Dim t = CreateDataStructure(poly1, poly2)
            Dim data1 = t.Item1
            Dim data2 = t.Item2

            Dim nbPtsInter = 0
            For Each vertex In data1
                If vertex.Intersect Then
                    nbPtsInter += 1
                End If
            Next

            If nbPtsInter = 0 Then
                'Vérification d'imbrication
                If data1(0).IsVertexInsidePolygonVertices(data2) Then
                    result = Nothing
                ElseIf data2(0).IsVertexInsidePolygonVertices(data1) Then
                    result.Add(poly1)
                    result.Add(poly2)
                Else
                    result.Add(poly1)
                End If
            ElseIf nbPtsInter Mod 2 = 0 Then
                Dim controlList As New List(Of VertexData)
                Dim index = 0
                While index < data1.Count AndAlso controlList.Count < data1.Count
                    If (data1(index).Intersect OrElse Not data1(index).IsVertexInsidePolygonVertices(data2)) AndAlso Not controlList.Contains(data1(index)) Then
                        Dim currentList As New List(Of Point)
                        Dim first = data1(index)
                        currentList.Add(first.Point)
                        Dim current = first
                        If Not controlList.Contains(current) Then
                            controlList.Add(current)
                        End If
                        'Dim isData1 = True
                        Do
                            If current.Intersect AndAlso current.EntryExit Then
                                current = current.Neighbour
                                Do
                                    current = current.Suiv
                                    If current <> first Then
                                        currentList.Add(current.Point)
                                        If currentList.Count > Math.Max(data1.Count, data2.Count) Then
                                            Throw New OverflowException("Résultat dépassant le nombre maximum de points possibles")
                                        End If
                                    End If
                                Loop Until current.Intersect
                                current = current.Neighbour
                                'If current <> first Then
                                '    index = Data1.IndexOf(current)
                                'End If
                            Else
                                current = current.Suiv
                                If current <> first Then
                                    currentList.Add(current.Point)
                                    If currentList.Count > Math.Max(data1.Count, data2.Count) Then
                                        Throw New OverflowException("Résultat dépassant le nombre maximum de points possibles")
                                    End If
                                    If Not controlList.Contains(current) Then
                                        controlList.Add(current)
                                    End If
                                End If
                                index += 1
                            End If
                        Loop Until current = first
                        If currentList.Count >= 3 Then
                            result.Add(New Surface(currentList, Surface.SurfaceType.Polygon).CleanPoints())
                        End If
                    Else
                        If Not controlList.Contains(data1(index)) Then
                            controlList.Add(data1(index))
                        End If
                    End If
                    index += 1
                End While
            End If
        Else
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If
        Return result
    End Function

    Public Function SurfacePolyBoolDifference(poly1 As Surface, poly2 As Surface, Optional tolerance As Integer? = Nothing) As List(Of Surface)
        Dim tolerance2 As Integer
        If tolerance Is Nothing Then
            tolerance2 = DefaultTolerance
        Else
            tolerance2 = tolerance
        End If

        Dim result As New List(Of Surface)
        If poly1.Points.Count = 0 OrElse poly2.Points.Count = 0 Then
            Throw New Exception("Empty surface(s)")
            GoTo Vide
        End If
        If Not poly1.Type = Surface.SurfaceType.Polygon OrElse Not poly2.Type = Surface.SurfaceType.Polygon Then
            Throw New NotImplementedException("Sorry only polygons can be intersected.")
        End If

        Dim polyOp = New PolyBool

        Dim polyL1 As New PointList From {}
        Dim polyL2 As New PointList From {}

        For Each pt In poly1.Points
            Dim ppt = New PolyBoolCS.Point(Math.Round(pt.X, -tolerance2), Math.Round(pt.Y, -tolerance2))
            polyL1.Add(ppt)
        Next
        For Each pt In poly2.Points
            Dim ppt = New PolyBoolCS.Point(Math.Round(pt.X, -tolerance2), Math.Round(pt.Y, -tolerance2))
            polyL2.Add(ppt)
        Next
        Dim polygon1 = New Polygon
        polygon1.regions = New List(Of PointList)() From {polyL1}
        Dim polygon2 = New Polygon
        polygon2.regions = New List(Of PointList)() From {polyL2}

        Dim polygonDifference = polyOp.difference(polygon1, polygon2)

        For Each inter In polygonDifference.regions
            Dim currList As New List(Of Point) From {}
            For Each ppt In inter
                currList.Add(New Point(ppt.x, ppt.y))
            Next
            result.Add(New Surface(currList, Surface.SurfaceType.Polygon))
        Next
Vide:
        Return result
    End Function
#End Region


    ''' <summary>
    ''' Line elongation.
    ''' </summary>
    ''' <param name="startPoint"> First Point of the line </param>
    ''' <param name="endPoint"> Second Point of the line </param>
    ''' <param name="startOffset"> Offset for first Point </param>
    ''' <param name="endOffset"> Offset for second Point </param>
    ''' <returns> New Line elongated </returns>
    <Obsolete("See Line.ExtendStart and Line.ExtendEnd")>
    Public Function Extrapol(startPoint As Point, endPoint As Point, startOffset As Double, endOffset As Double) As Line
        Dim directorVector = New Vector(startPoint, endPoint)
        Dim newStart = directorVector.Reverse.GetPointAtDistance(startOffset, startPoint)
        Dim newEnd = directorVector.GetPointAtDistance(endOffset, endPoint)
        Return New Line(newStart, newEnd)
    End Function

    ''' <summary>
    ''' Gravity Center of Points with optional weights.
    ''' </summary>
    ''' <param name="pointsList"></param>
    ''' <param name="weightsList"></param>
    ''' <returns></returns>
    Public Function Barycenter(pointsList As List(Of Point), Optional weightsList As List(Of Double) = Nothing) As Point
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
        Return New Point(baryX, baryY, baryZ)

    End Function


End Module

Module NoOverflows
    Public Function LongToInteger(ByVal value As Long) As Integer
        Dim cast As Caster
        cast.LongValue = value
        Return cast.IntValue
    End Function

    <StructLayout(LayoutKind.Explicit)>
    Private Structure Caster
        <FieldOffset(0)> Public LongValue As Long
        <FieldOffset(0)> Public IntValue As Integer
    End Structure
End Module
