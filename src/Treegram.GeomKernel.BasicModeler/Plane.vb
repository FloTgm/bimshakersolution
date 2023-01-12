''' <summary>
''' Represents a 3D-Plane
''' </summary>
Public Class Plane
    ''' <summary>
    ''' Normal Vector of the Plan
    ''' </summary>
    Public Vector As Vector

    ''' <summary>
    ''' Origin of the Plan
    ''' </summary>
    Public Origin As Point

    <Obsolete>
    Public Shared ApproximateCalculus = False
    <Obsolete>
    Public Shared Precision = 10

    'Cartesian equation of the Plan for a Point(x,y,z) : A*x + B*y + C*z + D = 0

    Private _a As Double?
    ''' <summary>
    ''' Get x-coefficient of Plan cartesian equation : A*x + B*y + C*z + D = 0
    ''' </summary>
    ''' <returns></returns>
    Public Property A As Double
        Set(value As Double)
            _a = Value
        End Set
        Get
            If _a Is Nothing Then
                _a = vector.X
            End If
            Return _a
        End Get
    End Property

    Private _b As Double?
    ''' <summary>
    ''' Get y-coefficient of Plan cartesian equation : A*x + B*y + C*z + D = 0
    ''' </summary>
    ''' <returns></returns>
    Public Property B As Double
        Set(value As Double)
            _b = Value
        End Set
        Get
            If _b Is Nothing Then
                _b = vector.Y
            End If
            Return _b
        End Get
    End Property

    Private _c As double?
    ''' <summary>
    ''' Get z-coefficient of Plan cartesian equation : A*x + B*y + C*z + D = 0
    ''' </summary>
    ''' <returns></returns>
    Public  Property C As Double
        Set(value As Double)
            _c = Value
        End Set
        Get
            If _c Is Nothing Then
                _c = vector.Z
            End If
            Return _c
        End Get
    End Property

    Private _d As Double?
    ''' <summary>
    ''' Get constant-coefficient of Plan cartesian equation : A*x + B*y + C*z + D = 0
    ''' </summary>
    ''' <returns></returns>
    Public Property D As Double
        Set(value As Double)
            _d = value
        End Set
        Get
            If _d Is Nothing Then
                _d = RoundToSignificantDigits(-(A * Origin.X + B * Origin.Y + C * Origin.Z), RealPrecision)
            End If
            Return _d
        End Get
    End Property

    ''' <summary>
    ''' Create new Plane by origin and vector.
    ''' </summary>
    ''' <param name="origin"> Origin of Plane </param>
    ''' <param name="vector"> Normal Vector of plane </param>
    Public Sub New(origin As Point, vector As Vector)
        Me.Origin = origin
        Me.Vector = vector
    End Sub

    ''' <summary>
    ''' Create new Plane by three points.
    ''' </summary>
    ''' <param name="ptA"></param>
    ''' <param name="ptB"></param>
    ''' <param name="ptC"></param>
    Public Sub New(ptA As Point, ptB As Point, ptC As Point)
        Dim normal = Vector.CrossProduct(New Vector(ptA, ptB), New Vector(ptA, ptC))
        If normal.Length = 0 Then
            Throw New ArgumentException("Points are aligned or confused.")
        End If
        Me.Origin = ptA
        Me.Vector = normal
    End Sub

    ''' <summary>
    ''' Check if a point belongs to the Plane
    ''' </summary>
    ''' <param name="point"> Cheched point. </param>
    ''' <param name="tolerance">Set the Tolerance as 10^-tolerance. Default is -4. tolerance = maximal distance to plane.</param>
    ''' <returns></returns>
    Public Function DoesContainPoint(point As Point, Optional tolerance As Integer? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = DefaultTolerance
        Dim result As Double
        result = RoundToSignificantDigits(A * point.X + B * point.Y + C * point.Z + D, RealPrecision)
        Return Math.Abs(result) <= Math.Pow(10, tolerance)
    End Function

    Public Overrides Function GetHashCode() As Integer
        return Tuple.Create(A,B,C,D).GetHashCode()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        If obj Is Nothing Then
            Throw New NullReferenceException
        ElseIf obj.GetType Is GetType(Plane) Then
            Return Me = obj
        Else
            Return False
        End If
    End Function

    Public Shared Operator =(p1 As Plane, p2 As Plane) As Boolean
        If p1 Is Nothing And p2 Is Nothing Then
            Return True
        ElseIf p1 Is Nothing And p2 IsNot Nothing Or p1 IsNot Nothing And p2 Is Nothing Then
            Throw New NullReferenceException
        Else
            Return p1.GetHashCode() = p2.GetHashCode()
            'Return (p1.DoesContainPoint(p2.Origin) AndAlso (p1.Vector.ScaleOneVector = p2.Vector.ScaleOneVector OrElse p1.Vector.ScaleOneVector = p2.Vector.Reverse.ScaleOneVector))
        End If
    End Operator

    Public Shared Operator <>(p1 As Plane, p2 As Plane) As Boolean
        If p1 Is Nothing And p2 Is Nothing Then
            Throw New NullReferenceException
        ElseIf p1 Is Nothing And p2 IsNot Nothing Or p1 IsNot Nothing And p2 Is Nothing Then
            Return True
        Else
            Return Not (p1 = p2)
        End If
    End Operator


    Public Overrides Function ToString() As String
        Return Origin.ToString + " -> " + Vector.ToString
    End Function

    ''' <summary>
    ''' Orthogonal projection of a point on the Plane
    ''' </summary>
    ''' <param name="pt"> Point to project </param>
    ''' <returns> Projected Point </returns>
    Public Function Projection(pt As Point) As Point
        If Vector.Length = 0 Then
            Return pt
        Else
            Dim k = ((Origin.X - pt.X) * Vector.X + (Origin.Y - pt.Y) * Vector.Y + (Origin.Z - pt.Z) * Vector.Z) / (Vector.InternLength * Vector.InternLength)
            Return New Point() With {
            .X = pt.X + k * Vector.X,
            .Y = pt.Y + k * Vector.Y,
            .Z = pt.Z + k * Vector.Z
            }
        End If
    End Function
End Class
