''' <summary>
''' Represents a 2D/3D-Vector
''' </summary>
Public Class Vector

    Private _vx, _vy, _vz As Double

    ''' <summary>
    ''' Coordinate X
    ''' </summary>
    ''' <returns></returns>
    Public Property X As Double
        Get
            Return _vx
        End Get
        Set(value As Double)
            _vx = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property
    ''' <summary>
    ''' Coordinate Y
    ''' </summary>
    ''' <returns></returns>
    Public Property Y As Double
        Get
            Return _vy
        End Get
        Set(value As Double)
            _vy = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property
    ''' <summary>
    ''' Coordinate Z
    ''' </summary>
    ''' <returns></returns>
    Public Property Z As Double
        Get
            Return _vz
        End Get
        Set(value As Double)
            _vz = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property

    ''' <summary>
    ''' Origin of the Vector
    ''' </summary>
    ''' <returns></returns>
    Public Property Origin As Point = {0.0, 0.0, 0.0}

    ''' <summary>
    ''' Get or Set EndPoint of the Vector
    ''' </summary>
    ''' <returns></returns>
    Public Property EndPoint As Point
        Get
            Return New Point(Origin.X + Me.X, Origin.Y + Me.Y, Origin.Z + Me.Z)
        End Get
        Set(value As Point)
            Me.X = value.X - Origin.X
            Me.Y = value.Y - Origin.Y
            Me.Z = value.Z - Origin.Z
        End Set
    End Property

    ''' <summary>
    ''' Get the Length of the Vector
    ''' </summary>
    ''' <returns></returns>
    <Obsolete("Use the Length Property")>
    Public ReadOnly Property Distance As Double
        Get
            Return Math.Round(RoundToSignificantDigits(Math.Sqrt(X * X + Y * Y + Z * Z), RealPrecision), -DefaultTolerance)
        End Get
    End Property

    ''' <summary>
    ''' Get the Length of the Vector
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Length As Double
        Get
            Return Math.Round(RoundToSignificantDigits(Math.Sqrt(X * X + Y * Y + Z * Z), RealPrecision), -DefaultTolerance)
        End Get
    End Property

    Friend ReadOnly Property InternLength As Double
        Get
            Return RoundToSignificantDigits(Math.Sqrt(X * X + Y * Y + Z * Z), RealPrecision)
        End Get
    End Property

    <Obsolete>
    Public Shared ApproximateCalculus = False
    <Obsolete>
    Public Shared Precision = 10

    ''' <summary>
    ''' Get the corrected standard Vector
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ScaleOneVector As Vector
        Get
            'If ApproximateCalculus Then
            '    Return New Vector With {.X = Math.Round(ProperX, Precision), .Y = Math.Round(ProperY, Precision), .Z = Math.Round(ProperZ, Precision)}
            'Else
            Return New Vector With {.X = ProperX, .Y = ProperY, .Z = ProperZ}
            'End If
        End Get
    End Property

    ''' <summary>
    ''' Create new null Vector : (0,0,0)
    ''' </summary>
    Public Sub New()
        Me.Origin = New Point
        Me.X = 0.0
        Me.Y = 0.0
        Me.Z = 0.0
    End Sub

    ''' <summary>
    ''' Create new Vector based on X Y Z
    ''' </summary>
    Public Sub New(x As Double, y As Double, z As Double)
        Me.Origin = New Point
        Me.X = x
        Me.Y = y
        Me.Z = z
    End Sub

    ''' <summary>
    ''' Create a new vector based on two Points
    ''' </summary>
    ''' <param name="pointA"> Origin Point</param>
    ''' <param name="pointB"> EndPoint </param>
    Public Sub New(pointA As Point, pointB As Point)
        Me.X = pointB.X - pointA.X
        Me.Y = pointB.Y - pointA.Y
        Me.Z = pointB.Z - pointA.Z
        Me.Origin = pointA
    End Sub

    ''' <summary>
    ''' Get the first coordinate of the normalized Vector
    ''' </summary>
    ''' <returns> ProperX </returns>
    Public ReadOnly Property ProperX As Double
        Get
            If Length = 0.0 Then
                Return 0
            Else
                Return RoundToSignificantDigits(X / InternLength, RealPrecision)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Get the second coordinate of the normalized Vector
    ''' </summary>
    ''' <returns> ProperY </returns>
    Public ReadOnly Property ProperY As Double
        Get
            If Length = 0.0 Then
                Return 0
            Else
                Return RoundToSignificantDigits(Y / InternLength, RealPrecision)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Get the third coordinate of the normalized Vector
    ''' </summary>
    ''' <returns> properZ </returns>
    Public ReadOnly Property ProperZ As Double
        Get
            If Length = 0.0 Then
                Return 0
            Else
                Return RoundToSignificantDigits(Z / InternLength, RealPrecision)
            End If

        End Get
    End Property

    Public Overloads Function ToString() As String
        If Length = 0.0 Then
            Return "{" & X & ", " & Y & ", " & Z & "} with a scale of " & Length
        End If
        If Length <> ScaleOneVector.Length And Length <> 1 Then
            Return "{" & X & ", " & Y & ", " & Z & "} with a scale of " & Length & Chr(10) & ScaleOneVector.ToString
        Else
            Return "{" & X & ", " & Y & ", " & Z & "} with a scale of " & Length
        End If
    End Function

    ''' <summary>
    ''' Get a Point by translating the Point origin following the Vector
    ''' </summary>
    ''' <param name="distance"> Distance of translation </param>
    ''' <param name="origin"> Origin of translation, by default Point(0,0,0) </param>
    ''' <returns> New Point translated </returns>
    <Obsolete("Take GetPointAtDistance")>
    Public Function GetPointOnVector(distance As Double, Optional origin As Point = Nothing) As Point
        If origin Is Nothing Then
            origin = New Point()
        End If
        Return New Point With {.X = origin.X + ProperX * distance, .Y = origin.Y + ProperY * distance, .Z = origin.Z + ProperZ * distance}
    End Function

    ''' <summary>
    ''' Get a Point by translating the Point origin following the Vector
    ''' </summary>
    ''' <param name="distance"> Distance of translation </param>
    ''' <param name="origin"> Origin of translation, by default Point(0,0,0) </param>
    ''' <returns> New Point translated </returns>
    Public Function GetPointAtDistance(distance As Double, Optional origin As Point = Nothing) As Point
        If origin Is Nothing Then
            origin = New Point()
        End If
        Return New Point With {.X = origin.X + ProperX * distance, .Y = origin.Y + ProperY * distance, .Z = origin.Z + ProperZ * distance}
    End Function

    ''' <summary>
    ''' Get a Point at Parameter t. (0 -> Vector.Origin, 1 -> Vector.EndPoint)
    ''' </summary>
    ''' <param name="t"> parameter </param>
    ''' <returns></returns>
    Public Function GetPointAtParameter(t As Double) As Point
        If t = 1 Then
            Return Me.EndPoint
        End If
        Return New Point(Me.Origin.X + t / Me.InternLength * Me.X, Me.Origin.Y + t / Me.InternLength * Me.Y, Me.Origin.Z + t / Me.InternLength * Me.Z)
    End Function

    ''' <summary>
    ''' Reverse Vector direction
    ''' </summary>
    ''' <returns> Reverted Vector </returns>
    Public Function Reverse() As Vector
        Return New Vector() With {.X = -Me.X, .Y = -Me.Y, .Z = -Me.Z, .Origin = Me.Origin}
    End Function


    Public Overrides Function GetHashCode() As Integer
        Return Tuple.Create(X, Y, Z).GetHashCode()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return Me = obj
    End Function

    Public Shared Operator =(v1 As Vector, v2 As Vector) As Boolean
        If IsNothing(v1) AndAlso Not IsNothing(v2) OrElse Not IsNothing(v1) AndAlso IsNothing(v2) Then
            'One is nothing, the other not
            Return False
        ElseIf IsNothing(v1) AndAlso IsNothing(v2) Then
            ' Both are nothing
            Return True
        Else
            Return v1.X = v2.X AndAlso v1.Y = v2.Y AndAlso v1.Z = v2.Z
        End If
    End Operator

    Public Shared Operator <>(v1 As Vector, v2 As Vector) As Boolean
        If IsNothing(v1) AndAlso Not IsNothing(v2) OrElse Not IsNothing(v1) AndAlso IsNothing(v2) Then
            Return True
        ElseIf IsNothing(v1) And IsNothing(v2) Then
            Throw New NullReferenceException
        Else
            Return Not (v1 = v2)
        End If
    End Operator

    Public Shared Operator *(v1 As Vector, v2 As Vector) As Double
        If IsNothing(v1) Or IsNothing(v2) Then
            Throw New NullReferenceException
        Else
            Return RoundToSignificantDigits(v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z, RealPrecision)
        End If
    End Operator

    Public Shared Operator *(v1 As Vector, facteur As Double) As Vector
        If IsNothing(v1) Then
            Throw New NullReferenceException
        Else
            Return New Vector() With {.X = v1.X * facteur, .Y = v1.Y * facteur, .Z = v1.Z * facteur}
        End If
    End Operator

    Public Shared Operator *(facteur As Double, v1 As Vector) As Vector
        If IsNothing(v1) Then
            Throw New NullReferenceException
        Else
            Return New Vector() With {.X = v1.X * facteur, .Y = v1.Y * facteur, .Z = v1.Z * facteur}
        End If
    End Operator

    Public Shared Operator +(v1 As Vector, v2 As Vector) As Vector
        If IsNothing(v1) Or IsNothing(v2) Then
            Throw New NullReferenceException
        Else
            Return New Vector With {
                .X = v1.X + v2.X,
                .Y = v1.Y + v2.Y,
                .Z = v1.Z + v2.Z
            }
        End If
    End Operator

    Public Shared Operator -(v1 As Vector, v2 As Vector) As Vector
        If IsNothing(v1) Or IsNothing(v2) Then
            Throw New NullReferenceException
        Else
            Return New Vector With {
                .X = v1.X - v2.X,
                .Y = v1.Y - v2.Y,
                .Z = v1.Z - v2.Z
            }
        End If
    End Operator

    ''' <summary>
    ''' Oriented angle compared to V2 in 2D.
    ''' For example, if V1 is Yaxis(0.1.0) and V2 is Xaxis(1.0.0). V1.Angle(V2,True) will be 90°.
    ''' </summary>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="valeurEnDegre"> Condition for unit </param>
    ''' <returns> Angle value in radians or degrees </returns>
    Function Angle(v2 As Vector, Optional valeurEnDegre As Boolean = False) As Double
        'If ApproximateCalculus Then

        '    If Not valeurEnDegre Then
        '        Return Math.Round(DecimalMath.DecimalMath.Acos(Math.Round(CDec((Me * V2)) / CDec((Me.Distance * V2.Distance)), 10)), Math.Min(Precision, 10)) * Sign(CrossProduct(V2, Me).Z)
        '    Else
        '        Return Math.Round((DecimalMath.DecimalMath.Acos(Math.Round(CDec((Me * V2)) / CDec((Me.Distance * V2.Distance)), 10))) * 180 / Math.PI, Math.Min(Precision, 10)) * Sign(CrossProduct(V2, Me).Z)
        '    End If
        'Else
        Dim denominateur = CDec(DecimalMath.DecimalMath.Sqrt(Me.X * Me.X + Me.Y * Me.Y + Me.Z * Me.Z) * DecimalMath.DecimalMath.Sqrt(v2.X * v2.X + v2.Y * v2.Y + v2.Z * v2.Z))
        If denominateur = 0 Then
            Return 0
        End If
        Dim x As Decimal = CDec(Me.X * v2.X + Me.Y * v2.Y + Me.Z * v2.Z) / CDec(DecimalMath.DecimalMath.Sqrt(Me.X * Me.X + Me.Y * Me.Y + Me.Z * Me.Z) * DecimalMath.DecimalMath.Sqrt(v2.X * v2.X + v2.Y * v2.Y + v2.Z * v2.Z))
        If x > 1 Then x = 1
        If x < -1 Then x = -1
        Dim aCos As Decimal = DecimalMath.DecimalMath.Acos(x)
        If Not valeurEnDegre Then
            Return Math.Round(RoundToSignificantDigits(aCos * Sign(CrossProduct(v2, Me).Z), RealPrecision), -DefaultTolerance)
        Else
            Return Math.Round(RoundToSignificantDigits((aCos * 180 / DecimalMath.DecimalMath.Pi) * Sign(CrossProduct(v2, Me).Z), RealPrecision), -DefaultTolerance)
        End If

        'End If
    End Function

    Private Function Sign(doble As Double) As Integer
        If doble >= 0 Then
            Return 1
        Else
            Return -1
        End If
    End Function

    ''' <summary>
    ''' Prefer function Rotate2.
    ''' </summary>
    ''' <param name="angleInRadian"> Angle in radians </param>
    <Obsolete("Use the function Rotate2.")>
    Public Sub Rotate(angleInRadian As Double)
        Dim endPoint = Me.EndPoint
        endPoint.X = Me.Origin.X + (Me.EndPoint.X - Me.Origin.X) * DecimalMath.DecimalMath.Cos(angleInRadian) - (Me.EndPoint.Y - Me.Origin.Y) * DecimalMath.DecimalMath.Sin(angleInRadian)
        endPoint.Y = Me.Origin.Y + (Me.EndPoint.X - Me.Origin.X) * DecimalMath.DecimalMath.Sin(angleInRadian) + (Me.EndPoint.Y - Me.Origin.Y) * DecimalMath.DecimalMath.Cos(angleInRadian)
        Me.EndPoint = endPoint
    End Sub

    ''' <summary>
    ''' Create new rotated Vector around the Z axis from OriginPoint
    ''' </summary>
    ''' <param name="angleInRadian"> Angle in radians </param>
    Public Function Rotate2(angleInRadian As Double) As Vector
        Return New Vector With {
            .Origin = Me.Origin,
            .EndPoint = New Point() With {
                .X = Me.Origin.X + (Me.EndPoint.X - Me.Origin.X) * DecimalMath.DecimalMath.Cos(angleInRadian) - (Me.EndPoint.Y - Me.Origin.Y) * DecimalMath.DecimalMath.Sin(angleInRadian),
                .Y = Me.Origin.Y + (Me.EndPoint.X - Me.Origin.X) * DecimalMath.DecimalMath.Sin(angleInRadian) + (Me.EndPoint.Y - Me.Origin.Y) * DecimalMath.DecimalMath.Cos(angleInRadian),
                .Z = Me.EndPoint.Z
            }
        }
    End Function

    Public Shared Narrowing Operator CType(v As Point) As Vector
        Return New Vector With {.X = v.X, .Y = v.Y, .Z = v.Z}
    End Operator

    ''' <summary>
    ''' Cross Product of two Vectors
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <returns> New Vector orthogonal to input Vectors </returns>
    Public Shared Function CrossProduct(v1 As Vector, v2 As Vector) As Vector
        Return New Vector With {.X = v1.Y * v2.Z - v1.Z * v2.Y, .Y = v1.Z * v2.X - v1.X * v2.Z, .Z = v1.X * v2.Y - v1.Y * v2.X}
    End Function

#Region "Alternative function for scalar product"
    'Public Shared Function ScalarProduct(v1 As Vector, v2 As Vector) As Double
    '    If ApproximateCalculus Then
    '        Return Math.Round(v1.Distance * v2.Distance * Math.Cos(v1.Angle(v2)), Precision)
    '    Else
    '        Return v1.Distance * v2.Distance * Math.Cos(v1.Angle(v2))
    '    End If
    'End Function
#End Region

    ''' <summary>
    ''' Check parallelism of Vectors
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance"> Set the Tolerance as 10^-tolerance (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
    ''' <returns> Parallelism of two Vectors </returns>
    <Obsolete("Use the function AreCollinear")>
    Public Shared Function AreVectorsParallel(v1 As Vector, v2 As Vector, Optional tolerance As Integer? = Nothing) As Boolean ''transformer en tolerance d'angle p ê
        If tolerance Is Nothing Then tolerance = -DefaultTolerance
        Dim s1 = v1.ScaleOneVector
        Dim s2 = v2.ScaleOneVector
        s1.Round(tolerance)
        s2.Round(tolerance)
        Return s1 = s2 OrElse s1 = s2.Reverse()
    End Function

    ''' <summary>
    ''' Check in 3D if two vectors are collinear.
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance"> Set the Tolerance as a power of 10 (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
    ''' <returns> Collinearity of two Vectors.</returns>
    Public Shared Function AreCollinear(v1 As Vector, v2 As Vector, Optional tolerance As Integer? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = DefaultTolerance
        Dim s1 = v1.ScaleOneVector
        Dim s2 = v2.ScaleOneVector
        s1.Round(-tolerance)
        s2.Round(-tolerance)
        Return s1 = s2 OrElse s1 = s2.Reverse()
    End Function

    ''' <summary>
    ''' Check in 3D if two vectors are collinear.
    ''' </summary>
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance"> Set the Tolerance (evaluate the vectors proper coordinates), 10^DefaultTolerance is used by default. </param>
    ''' <returns> Collinearity of two Vectors.</returns>
    Public Shared Function AreCollinear(v1 As Vector, v2 As Vector, Optional tolerance As Double? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = Math.Pow(10, DefaultTolerance)
        Return (Math.Abs(v1.ProperX - v2.ProperX) <= tolerance AndAlso Math.Abs(v1.ProperY - v2.ProperY) <= tolerance AndAlso Math.Abs(v1.ProperZ - v2.ProperZ) <= tolerance) OrElse
            (Math.Abs(v1.ProperX + v2.ProperX) <= tolerance AndAlso Math.Abs(v1.ProperY + v2.ProperY) <= tolerance AndAlso Math.Abs(v1.ProperZ + v2.ProperZ) <= tolerance)
    End Function

    ''' <summary>
    ''' Check in 3D if two vectors = segments are aligned on the exact same line.
    ''' </summary>
    ''' It's not a check of parallelism
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance">Set the Tolerance as power of 10 (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
    ''' <returns> Result of check as Boolean </returns>
    Public Shared Function AreAligned(v1 As Vector, v2 As Vector, Optional tolerance As Integer? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = DefaultTolerance
        If AreCollinear(v1, v2, tolerance) Then
            If Application.Distance(v1.Origin, v2, True) <= 1 * Math.Pow(10, tolerance) AndAlso Application.Distance(v1.EndPoint, v2, True) <= 1 * Math.Pow(10, tolerance) Then
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Check in 3D if two vectors = segments are on the exact same line
    ''' </summary>
    ''' It's not a check of parallelism
    ''' <param name="v1"> First Vector </param>
    ''' <param name="v2"> Second Vector </param>
    ''' <param name="tolerance">Set the Tolerance (evaluate the vectors proper coordinates), 10^DefaultTolerance is used by default. </param>
    ''' <returns> Result of check as Boolean </returns>
    Public Shared Function AreAligned(v1 As Vector, v2 As Vector, Optional tolerance As Double? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = Math.Pow(10, DefaultTolerance)
        If AreCollinear(v1, v2, tolerance) Then
            If Application.Distance(v1.Origin, v2, True) <= tolerance AndAlso Application.Distance(v1.EndPoint, v2, True) <= tolerance Then
                Return True
            End If
        End If
        Return False
    End Function

    ''' <summary>
    ''' Rounds Vector coordinates to a specified number of fractional digits. The maximum is RealPrecision, it is the default precision.
    ''' </summary>
    ''' <param name="digits"> Number of fractional digits of precision </param>
    Public Sub Round(digits As Integer)
        If digits < RealPrecision Then
            Me.X = Math.Round(X, digits)
            Me.Y = Math.Round(Y, digits)
            Me.Z = Math.Round(Z, digits)
        End If
    End Sub

    Public Shared Function Determinant(v1 As Vector, v2 As Vector, v3 As Vector) As Double
        Dim det = v1.X * v2.Y * v3.Z + v1.Y * v2.Z * v3.X + v1.Z * v2.X * v3.Y - v1.Z * v2.Y * v3.X - v1.X * v2.Z * v3.Y - v1.Y * v2.X * v3.Z
        Return RoundToSignificantDigits(det, RealPrecision)
    End Function

    Public Shared Function DeterminantZ(v1 As Vector, v2 As Vector) As Double
        Return RoundToSignificantDigits(v1.X * v2.Y - v1.Y * v2.X, RealPrecision)
    End Function

#Region "Not yet implemented"
    Public Function IsVectorInFrontOfThis(v As Vector) As Boolean
        Throw New NotImplementedException
    End Function
#End Region

End Class
