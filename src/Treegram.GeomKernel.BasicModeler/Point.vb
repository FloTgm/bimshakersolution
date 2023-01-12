''' <summary>
''' Represents a 2D/3D-Point
''' </summary>
Public Class Point

    Private _px, _py, _pz As Double

    ''' <summary>
    ''' X coordinate
    ''' </summary>
    ''' <returns></returns>
    Public Property X As Double
        Get
            Return _px
        End Get
        Set(value As Double)
            _px = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property

    ''' <summary>
    ''' Y coordinate
    ''' </summary>
    ''' <returns></returns>
    Public Property Y As Double
        Get
            Return _py
        End Get
        Set(value As Double)
            _py = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property

    ''' <summary>
    ''' Z coordinate
    ''' </summary>
    ''' <returns></returns>
    Public Property Z As Double
        Get
            Return _pz
        End Get
        Set(value As Double)
            _pz = RoundToSignificantDigits(value, RealPrecision)
        End Set
    End Property

    ''' <summary>
    ''' Create new Point : (0,0,0)
    ''' </summary>
    Public Sub New()
        Me.X = 0.0
        Me.Y = 0.0
        Me.Z = 0.0
    End Sub

    ''' <summary>
    ''' Create new Point by coordinates
    ''' </summary>
    ''' <param name="x"></param>
    ''' <param name="y"></param>
    ''' <param name="z"></param>
    Public Sub New(x As Double, y As Double, Optional z As Double = 0.0)
        Me.X = x
        Me.Y = y
        Me.Z = z
    End Sub

    Public Shared Narrowing Operator CType(v As Drawing.Point) As Point
        Return New Point With {.X = v.X, .Y = v.Y, .Z = 0.0}
    End Operator

    Public Shared Widening Operator CType(v As Point) As Drawing.Point
        Return New Drawing.Point With {.X = v.X, .Y = v.Y}
    End Operator

    Public Shared Narrowing Operator CType(v As Double()) As Point
        If v.Count = 2 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = 0.0}
        ElseIf v.Count = 3 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = v(2)}
        Else
            Throw New ArgumentException
        End If
    End Operator

    Public Shared Narrowing Operator CType(v As Integer()) As Point
        If v.Count = 2 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = 0}
        ElseIf v.Count = 3 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = v(2)}
        Else
            Throw New ArgumentException
        End If
    End Operator

    Public Shared Narrowing Operator CType(v As Object()) As Point
        If v.Count = 2 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = 0.0}
        ElseIf v.Count = 3 Then
            Return New Point With {.X = v(0), .Y = v(1), .Z = v(2)}
        Else
            Throw New ArgumentException
        End If
    End Operator

    Public Shared Widening Operator CType(v As Point) As Double()
        Return {v.X, v.Y, v.Z}
    End Operator

    Public Overrides Function ToString() As String
        Return $"({X}, {Y}, {Z})"
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return Me = obj
    End Function


    Public Overrides Function GetHashCode() As Integer
        Return Tuple.Create(X, Y, Z).GetHashCode()

        ''Fat
        'Return Me.ToString().GetHashCode()

        ''Not Working
        'Const hashingBase As Long = 2166136261
        'Const hashingMultiplier As Integer = 16777619
        'Dim hash = hashingBase
        'hash = UncheckLongToInteger((hash * hashingMultiplier) ^ X.GetHashCode())
        'hash = UncheckLongToInteger((hash * hashingMultiplier) ^ Y.GetHashCode())
        'hash = UncheckLongToInteger((hash * hashingMultiplier) ^ Z.GetHashCode())
        'Return hash
    End Function

    Private Shared Function UncheckLongToInteger(value As Long) As Integer
        Return value Mod Integer.MaxValue
    End Function

    Public Shared Operator =(pA As Point, pB As Point) As Boolean
        If pA Is Nothing And pB Is Nothing Then
            Return True
        ElseIf pA Is Nothing And Not pB Is Nothing Or Not pA Is Nothing And pB Is Nothing Then
            Return False
        Else
            Return pA.X = pB.X AndAlso pA.Y = pB.Y AndAlso pA.Z = pB.Z
        End If
    End Operator

    ''' <summary>
    ''' Equality check with a tolerance. 
    ''' </summary>
    ''' <param name="p1"></param>
    ''' <param name="p2"></param>
    ''' <param name="tolerance">represent the Math.Pow tolerance. For exemple, tolerance = -3 will check the equality with a Math.Pow(10, -3) precision</param>
    ''' <returns></returns>
    Public Shared Function AlmostEquals(p1 As Point, p2 As Point, Optional tolerance As Double? = Nothing) As Boolean
        If tolerance Is Nothing Then tolerance = Math.Pow(10, DefaultTolerance)
        Return (Math.Abs(p1.X - p2.X) <= tolerance AndAlso
                Math.Abs(p1.Y - p2.Y) <= tolerance AndAlso
                Math.Abs(p1.Z - p2.Z) <= tolerance)
    End Function

    Public Shared Operator *(p As Point, f As Double) As Point
        If (p Is Nothing) Then
            Return p
        Else
            Return {p.X * f, p.Y * f, p.Z * f}
        End If
    End Operator

    Public Shared Operator /(p As Point, f As Double) As Point
        If (p Is Nothing) Then
            Return p
        Else
            Return {p.X / f, p.Y / f, p.Z / f}
        End If
    End Operator

    Public Shared Operator +(pA As Point, pB As Point) As Point
        If (pA Is Nothing OrElse pB Is Nothing) Then
            Throw New ArgumentNullException()
        Else
            Return {pA.X + pB.X, pA.Y + pB.Y, pA.Z + pB.Z}
        End If
    End Operator

    Public Shared Operator -(pA As Point, pB As Point) As Point
        If (pA Is Nothing OrElse pB Is Nothing) Then
            Throw New ArgumentNullException()
        Else
            Return {pA.X - pB.X, pA.Y - pB.Y, pA.Z - pB.Z}
        End If
    End Operator

    Public Shared Operator <>(pA As Point, pB As Point) As Boolean
        If pA Is Nothing OrElse pB Is Nothing Then
            Return True
        Else
            Return Not (pA = pB)
        End If
    End Operator

    ''' <summary>
    ''' Rounds Point coordinates to a specified number of fractional digits.
    ''' </summary>
    ''' <param name="precision"> Number of wanted digits. </param>
    Public Sub Round(precision As Integer)
        'If Precision > RealPrecision Then
        Me.X = Math.Round(X, precision)
        Me.Y = Math.Round(Y, precision)
        Me.Z = Math.Round(Z, precision)
        'End If
    End Sub

    ''' <summary>
    ''' Prefer the function Translate2.
    ''' </summary>
    ''' <param name="vector"> Vector for translation </param>
    ''' <param name="distance"> Distance of translation </param>
    <Obsolete("Use the Function Translate2.")>
    Public Sub Translate(vector As Vector, distance As Double)
        Me.X = Me.X + vector.ProperX * distance
        Me.Y = Me.Y + vector.ProperY * distance
        Me.Z = Me.Z + vector.ProperZ * distance
    End Sub

    ''' <summary>
    ''' Create new translated Vector.
    ''' </summary>
    ''' <param name="vector">Vector for translation </param>
    ''' <param name="distance"> Distance of translation </param>
    ''' <returns></returns>
    Public Function Translate2(vector As Vector, distance As Double) As Point
        Return New Point() With {
            .X = Me.X + vector.ProperX * distance,
            .Y = Me.Y + vector.ProperY * distance,
            .Z = Me.Z + vector.ProperZ * distance
        }
    End Function

    ''' <summary>
    ''' Prefer the function Translate2.
    ''' </summary>
    ''' <param name="vector"> vector for translation </param>
    <Obsolete("Use the Function Translate2.")>
    Public Sub Translate(vector As Vector)
        Me.X = Me.X + vector.X
        Me.Y = Me.Y + vector.Y
        Me.Z = Me.Z + vector.Z
    End Sub

    ''' <summary>
    ''' Create new translated Vector.
    ''' </summary>
    ''' <param name="vector">Vector for translation </param>
    ''' <returns></returns>
    Public Function Translate2(vector As Vector) As Point
        Return New Point() With {
            .X = Me.X + vector.X,
            .Y = Me.Y + vector.Y,
            .Z = Me.Z + vector.Z
        }
    End Function

    Public Function Clone() As Point
        Return Me.MemberwiseClone
    End Function

End Class
