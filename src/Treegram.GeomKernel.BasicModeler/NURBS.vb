Public Class NurbScurve
    Private Property Points As List(Of Point)

    Private Property Weights As List(Of Double)

    Private Property Knots As List(Of Double)

    Private Property Degree As Integer

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="points"> Control points </param>
    ''' <param name="weights"> Weights of control points </param>
    ''' <param name="knots"> Knots of nodal vector (knots.count = points.count + degree + 1)  </param>
    ''' <param name="degree"> Degree of the form </param>
    Public Sub New(points As List(Of Point), weights As List(Of Double), knots As List(Of Double), degree As Integer)
        If points Is Nothing OrElse points.Count = 0 OrElse weights Is Nothing OrElse weights.Count = 0 OrElse knots Is Nothing OrElse knots.Count = 0 OrElse IsNothing(degree) Then
            Throw New NullReferenceException()
        End If
        If degree < 1 Then
            Throw New Exception("NURBS requires degree to be greater than 0")
        End If
        If points.Count < degree + 1 Then
            Throw New Exception("NURBS must have at least one more control point than degree")
        End If
        If points.Count <> weights.Count Then
            Throw New Exception("NURBS needs to have the same number of points and weights")
        End If
        If knots.Count <> points.Count + degree + 1 Then
            Throw New Exception("In NURBS the number of knots = number of points + degree + 1")
        End If
        If Not CorrectOrderKnots(knots) Then
            Throw New Exception("NURBS knots sequenbce should be non-decreasing")
        End If
        If Not CorrectMultiplicityKnots(knots, degree) Then
            Throw New Exception("Knot multiplicity is incorrect (degree+1 at the start/end, less than degree in the interior")
        End If
        'Autres cas d'exceptions
        '
        '
        '
        '
        Me.Points = points
        Me.Weights = Nothing
        For Each w In weights
            Me.Weights.Add(RoundToSignificantDigits(w, RealPrecision))
        Next
        Me.Knots = Nothing
        For Each k In knots
            Me.Knots.Add(RoundToSignificantDigits(k, RealPrecision))
        Next
        Me.Degree = degree
    End Sub

    ''' <summary>
    ''' Basic function for NURBS from Cox-De Boor algorithm
    ''' </summary>
    ''' <param name="j"> Index of the node [0,m-d-1-1] </param>
    ''' <param name="d"> NURBS degree </param>
    ''' <param name="t"> Curve parameter, between 0 and 1 </param>
    ''' <returns></returns>
    Private Function Njd(ByVal j As Integer, ByVal d As Integer, ByVal t As Double) As Double
        Dim normKnots = NormalizeKnots(Knots)
        Dim n As Double() = New Double(d + 1 - 1) {}
        Dim saved, temp As Double
        Dim m As Integer = normKnots.Count - 1
        If (j = 0 AndAlso t = normKnots(0)) OrElse (j = (m - d - 1) AndAlso t = normKnots(m)) Then Return 1.0R
        If t < normKnots(j) OrElse t >= normKnots(j + d + 1) Then Return 0R

        For i As Integer = 0 To d

            If t >= normKnots(j + i) AndAlso t < normKnots(j + i + 1) Then
                n(i) = 1.0R
            Else
                n(i) = 0R
            End If
        Next

        For k As Integer = 1 To d

            If n(0) = 0 Then
                saved = 0R
            Else
                saved = ((t - normKnots(j)) * n(0)) / (normKnots(j + k) - normKnots(j))
            End If

            For i As Integer = 0 To d - k + 1 - 1
                Dim uleft As Double = normKnots(j + i + 1)
                Dim uright As Double = normKnots(j + i + k + 1)

                If n(i + 1) = 0 Then
                    n(i) = saved
                    saved = 0R
                Else
                    temp = n(i + 1) / (uright - uleft)
                    n(i) = RoundToSignificantDigits(saved + (uright - t) * temp, RealPrecision)
                    saved = (t - uleft) * temp
                End If
            Next
        Next

        Return n(0)
    End Function

    ''' <summary>
    ''' Derives basic functions of NURBS
    ''' </summary>
    ''' <param name="j"></param>
    ''' <param name="d"></param>
    ''' <param name="t"></param>
    ''' <returns></returns>
    Private Function NjdPrime(j As Integer, d As Integer, t As Double) As Double
        Dim normKnots = NormalizeKnots(Knots)
        Dim a, b As Double
        If normKnots(j + d) > normKnots(j) Then
            a = d / (normKnots(j + d) - normKnots(j)) * Njd(j, d - 1, t)
        Else
            a = 0
        End If
        If normKnots(j + d + 1) > normKnots(j + 1) Then
            b = d / (normKnots(j + d + 1) - normKnots(j + 1)) * Njd(j + 1, d - 1, t)
        Else
            b = 0
        End If
        Return RoundToSignificantDigits(a - b, RealPrecision)
    End Function


    ''' <summary>
    ''' Curve position
    ''' </summary>
    ''' <param name="t"> Curve parameter between 0 and 1 </param>
    ''' <returns></returns>
    Public Function C(t As Double) As Point
        If t < 0 OrElse t > 1 Then
            Throw New Exception("t must be in the interval [0;1]")
        End If
        Dim newPoint As New Point
        Dim denum As Double = 0
        For i = 0 To Points.Count - 1
            newPoint.X += Weights(i) * Points(i).X * Njd(i, Degree, t)
            newPoint.Y += Weights(i) * Points(i).Y * Njd(i, Degree, t)
            newPoint.Z += Weights(i) * Points(i).Z * Njd(i, Degree, t)
            denum += Weights(i) * Njd(i, Degree, t)
        Next
        If denum <> 0 Then
            newPoint.X = newPoint.X / denum
            newPoint.Y = newPoint.Y / denum
            newPoint.Z = newPoint.Z / denum
        End If
        Return newPoint
    End Function

    ''' <summary>
    ''' Derived Curve position
    ''' </summary>
    ''' <param name="t"> Curve parameter between 0 and 1 </param>
    ''' <returns></returns>
    Public Function Cprime(t As Double) As Vector
        If t < 0 OrElse t > 1 Then
            Throw New Exception("t must be in the interval [0;1]")
        End If
        Dim newVector As New Vector
        Dim denum As Double = 0
        Dim ax, ay, az, b, cx, cy, cz, d As New Double
        For i = 0 To Knots.Count - Degree - 1 - 1
            ax += Weights(i) * Points(i).X * NjdPrime(i, Degree, t)
            ay += Weights(i) * Points(i).Y * NjdPrime(i, Degree, t)
            az += Weights(i) * Points(i).Z * NjdPrime(i, Degree, t)
            b += Weights(i) * Njd(i, Degree, t)
            cx = Weights(i) * Points(i).X * Njd(i, Degree, t)
            cy = Weights(i) * Points(i).Y * Njd(i, Degree, t)
            cz = Weights(i) * Points(i).Z * Njd(i, Degree, t)
            d = Weights(i) * NjdPrime(i, Degree, t)
        Next
        denum = b * b
        newVector.X = (ax * b - cx * d) / denum
        newVector.Y = (ay * b - cy * d) / denum
        newVector.Z = (az * b - cz * d) / denum
        Return newVector
    End Function

    Private Function CorrectOrderKnots(knots As List(Of Double)) As Boolean
        Dim current = knots.First
        For i = 1 To knots.Count - 1
            If current < 0 Then Return False
            If knots(i) < current Then Return False
            current = knots(i)
        Next
        If current = 0 Then Return False
        Return True
    End Function

    Private Function NormalizeKnots(knots As List(Of Double)) As List(Of Double)
        Dim newKnots As New List(Of Double)
        For Each node In knots
            newKnots.Add((node - knots.First) / (knots.Last - knots.First))
        Next
        Return newKnots
    End Function

    Private Function CorrectMultiplicityKnots(knots As List(Of Double), degree As Integer) As Boolean
        Dim dicoNodes As New Dictionary(Of Double, Integer)
        For Each knot In knots
            If dicoNodes.ContainsKey(knot) Then
                dicoNodes(knot) += 1
            Else
                dicoNodes.Add(knot, 1)
            End If
        Next
        If dicoNodes.First.Value <> degree + 1 OrElse dicoNodes.Last.Value <> degree + 1 Then
            Return False
        End If
        If dicoNodes.Count > 2 Then
            For i = 1 To dicoNodes.Count - 2
                If dicoNodes.ElementAt(i).Value > degree Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

    ''' <summary>
    ''' Approximate length of the NURBS curve by interval size
    ''' </summary>
    ''' <param name="digits"> Number of digits of the interval division </param>
    ''' <returns></returns>
    Private Function LengthByInterval(digits As Integer) As Double
        If digits > RealPrecision Then digits = RealPrecision
        If digits < 0 Then Return Nothing
        Dim prec As Double = Math.Pow(10, -digits)
        Dim distFinale = 0.0
        'Application.PRECISION = digits + 1
        If Degree > 1 Then
            Dim point1, point2 As Point
            point1 = Me.C(0)
            Dim t = 0.0
            While t < 1
                t += Math.Round(prec, digits)
                t = Math.Round(t, digits)
                point2 = Me.C(t)
                distFinale += Distance(point1, point2)
                point1 = point2
            End While
        Else
            For k = 0 To Points.Count - 2
                distFinale += Distance(Points(k), Points(k + 1))
            Next
        End If
        Return Math.Round(distFinale, digits)
    End Function

    ''' <summary>
    ''' Length of the controlPoints polygon
    ''' </summary>
    ''' <returns></returns>
    Private Function PrimaryLength() As Double
        Dim l = 0.0
        Dim r = 0
        For k = 0 To Points.Count - 2
            l += Distance(Points(k), Points(k + 1))
            r += 1
            If r = 4 Then
                l = RoundToSignificantDigits(l, RealPrecision)
                r = 0
            End If
        Next
        If r = 4 Then
            l = RoundToSignificantDigits(l, RealPrecision)
        End If
        Return l
    End Function

    ''' <summary>
    ''' Approximate Length
    ''' </summary>
    ''' <param name="digits"> Number of digits of precision </param>
    ''' <returns></returns>
    Public Function Length(digits As Integer) As Double
        Dim prec As Double = Math.Pow(10, -digits)
        Dim primary = PrimaryLength()
        Dim puissance = 0
        While Math.Pow(10, puissance) < primary / prec
            puissance += 1
        End While
        Return RoundToSignificantDigits(Math.Round(LengthByInterval(puissance + 1), digits), RealPrecision)
    End Function

    ''' <summary>
    ''' Approximate Length
    ''' </summary>
    ''' <param name="digits"> Number of digits of precision </param>
    ''' <returns></returns>
    Public Function Length2(digits As Integer) As Double
        Dim currentLength = PrimaryLength()
        Dim secondLength = LengthByInterval(1)
        Dim i = 1
        '(currentLength - secondLength) > Math.Pow(10, -digits - 1) OrElse
        While Math.Round(currentLength, digits) <> Math.Round(secondLength, digits)
            currentLength = secondLength
            secondLength = LengthByInterval(i + 1)
            i += 1
        End While
        Return RoundToSignificantDigits(Math.Round(currentLength, digits), RealPrecision)
    End Function

    Public Function IsClosed() As Boolean
        Return Points.First = Points.Last
    End Function
End Class
