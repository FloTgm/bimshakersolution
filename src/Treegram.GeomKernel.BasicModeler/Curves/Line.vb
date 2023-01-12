
''' <summary>
''' Represents a 2D/3D-Line (used like a segment or a infinite line following the case)
''' </summary>
Public Class Line : Inherits Curve

        <Obsolete>
        Public Shared ApproximateCalculus = False
        <Obsolete>
        Public Shared Precision = 10

    ''' <summary>
    ''' Vector of the Line
    ''' </summary>
    Private _mVector As Vector

    ''' <summary>
    ''' Get or Set the Vector of the Line
    ''' </summary>
    ''' <returns></returns>
    Public Property Vector As Vector
        Get
            Return _mVector
        End Get
        Set(vector As Vector)
            Me._mVector = vector
        End Set
        End Property

        ''' <summary>
        ''' Get or Set the StartPoint of the Line
        ''' </summary>
        ''' <returns> StartPoint </returns>
        Public Overrides Property StartPoint As Point
            Get
                Return Me.Vector.Origin
            End Get
            Set(startPoint As Point)
                Me.Vector.Origin = startPoint
            End Set
        End Property

        ''' <summary>
        ''' Get or Set the EndPoint of the Line
        ''' </summary>
        ''' <returns> EndPoint </returns>
        Public Overrides Property EndPoint As Point
            Get
                Return Me.Vector.EndPoint
            End Get
            Set(endPoint As Point)
                Me.Vector.EndPoint = endPoint
            End Set
        End Property

        ''' <summary>
        ''' Get the Length of the Line
        ''' </summary>
        ''' <returns> Length </returns>
        Public Overrides ReadOnly Property Length As Double
            Get
                Return Me.Vector.Length
            End Get
        End Property

        Friend ReadOnly Property InternLength As Double
            Get
                Return Me.Vector.InternLength
            End Get
        End Property

        ''' <summary>
        ''' Get the MidPoint of the Line
        ''' </summary>
        ''' <returns> MidPoint </returns>
        Public Overrides ReadOnly Property MidPoint As Point
            Get
                Return Application.MidPoint(StartPoint, EndPoint)
            End Get
        End Property

        ''' <summary>
        ''' Create new Line based on Vector
        ''' </summary>
        ''' <param name="vector"> Vector </param>
        Public Sub New(vector As Vector)
            Me.Vector = vector
        End Sub

        ''' <summary>
        ''' Create new Line based on two Points
        ''' </summary>
        ''' <param name="startPoint"> StartPoint </param>
        ''' <param name="endPoint"> EndPoint </param>
        Public Sub New(startPoint As Point, endPoint As Point)
            Me.Vector = New Vector(startPoint, endPoint)
        End Sub

        Public Shared Widening Operator CType(points As Point()) As Line
            If points.Count <> 2 Then
                Throw New ArgumentOutOfRangeException("points", "Nombre de points incorret.")
            ElseIf points(0) Is Nothing Or points(1) Is Nothing Then
                Throw New NullReferenceException("Au moins un des points est null.")
            End If
            Return New Line(points(0), points(1))
        End Operator

        Public Shared Widening Operator CType(vector As Vector) As Line
            Return New Line(vector)
        End Operator

        Public Shared Widening Operator CType(line As Line) As Vector
            Return New Vector(line.StartPoint, line.EndPoint)
        End Operator

        Public Shared Operator =(l1 As Line, l2 As Line) As Boolean
            If IsNothing(l1) OrElse IsNothing(l2) Then
                Return False
            Else
                'If ApproximateCalculus Then
                '    Return Math.Round(l1.StartPoint.X, Precision) = Math.Round(l2.StartPoint.X, Precision) And Math.Round(l1.StartPoint.Y, Precision) = Math.Round(l2.StartPoint.Y, Precision) And Math.Round(l1.StartPoint.Z, Precision) = Math.Round(l2.StartPoint.Z, Precision) And
                '         Math.Round(l1.EndPoint.X, Precision) = Math.Round(l2.EndPoint.X, Precision) And Math.Round(l1.EndPoint.Y, Precision) = Math.Round(l2.EndPoint.Y, Precision) And Math.Round(l1.EndPoint.Z, Precision) = Math.Round(l2.EndPoint.Z, Precision)
                'Else
                Return l1.StartPoint = l2.StartPoint AndAlso l1.EndPoint = l2.EndPoint
                'End If
            End If
        End Operator

        Public Shared Operator <>(l1 As Line, l2 As Line) As Boolean
            Return Not (l1 = l2)
        End Operator

        ''' <summary>
        ''' Equality check with a tolerance. Check startPoint and EndPoint.
        ''' </summary>
        ''' <param name="l1"></param>
        ''' <param name="l2"></param>
        ''' <param name="tolerance">represent the Math.Pow tolerance. For exemple, tolerance = -3 will check the equality with a Math.Pow(10, -3) precision</param>
        ''' <returns></returns>
        Public Shared Function AlmostEquals(l1 As Line, l2 As Line, Optional tolerance As Double? = Nothing) As Boolean
            If tolerance Is Nothing Then tolerance = Math.Pow(10, DefaultTolerance)
            Return (Math.Abs(l1.StartPoint.X - l2.StartPoint.X) <= tolerance AndAlso
                    Math.Abs(l1.StartPoint.Y - l2.StartPoint.Y) <= tolerance AndAlso
                    Math.Abs(l1.StartPoint.Z - l2.StartPoint.Z) <= tolerance AndAlso
                    Math.Abs(l1.EndPoint.X - l2.EndPoint.X) <= tolerance AndAlso
                    Math.Abs(l1.EndPoint.Y - l2.EndPoint.Y) <= tolerance AndAlso
                    Math.Abs(l1.EndPoint.Z - l2.EndPoint.Z) <= tolerance)
        End Function


        ''' <summary>
        ''' Reverse the Line direction.
        ''' </summary>
        ''' <returns> New Line </returns>
        Public Function Reverse() As Line
            Return New Line(Me.EndPoint, Me.StartPoint)
        End Function

        ''' <summary>
        ''' Extend the line from StartPoint.
        ''' </summary>
        ''' <param name="offset"> Algebric offset. </param>
        ''' <returns> New Line elongated. </returns>
        Public Function ExtendStart(offset As Double) As Line
            Dim newStPt = Me.Vector.GetPointAtDistance(-offset, Me.StartPoint)
            Return New Line(newStPt, Me.EndPoint)
        End Function
        ''' <summary>
        ''' Extend the Line from EndPoint.
        ''' </summary>
        ''' <param name="offset"> Algebric offset. </param>
        ''' <returns> New Line elongated. </returns>
        Public Function ExtendEnd(offset As Double) As Line
            Dim newEndPt = Me.Vector.GetPointAtDistance(offset, Me.EndPoint)
            Return New Line(Me.StartPoint, newEndPt)
        End Function


        ''' <summary>
        ''' Check in 3D if two lines = segments are aligned.
        ''' </summary>
        ''' It's not a check of parallelism
        ''' <param name="line1"> First Line </param>
        ''' <param name="line2"> Second Line </param>
        ''' <param name="tolerance">Set the Tolerance as power of 10 (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
        ''' <returns> Result of check as Boolean </returns>
        Public Shared Function AreAligned(line1 As Line, line2 As Line, Optional tolerance As Integer? = Nothing) As Boolean
            Return Vector.AreAligned(line1.Vector, line2.Vector, tolerance)
        End Function

        ''' <summary>
        ''' Check in 3D if two lines = segments are aligned.
        ''' </summary>
        ''' It's not a check of parallelism
        ''' <param name="line1"> First Line </param>
        ''' <param name="line2"> Second Line </param>
        ''' <param name="tolerance">Set the Tolerance (evaluate the vectors proper coordinates), 10^DefaultTolerance is used by default. </param>
        ''' <returns> Result of check as Boolean </returns>
        Public Shared Function AreAligned(line1 As Line, line2 As Line, Optional tolerance As Double? = Nothing) As Boolean
            Return Vector.AreAligned(line1.Vector, line2.Vector, tolerance)
        End Function

        ''' <summary>
        ''' Check in 3D if two lines = segments are parallel.
        ''' </summary>
        ''' It's not a check of parallelism
        ''' <param name="line1"> First Line </param>
        ''' <param name="line2"> Second Line </param>
        ''' <param name="tolerance">Set the Tolerance as power of 10 (evaluate the vectors proper coordinates), DefaultTolerance is used by default. </param>
        ''' <returns> Result of check as Boolean </returns>
        Public Shared Function AreParallel(line1 As Line, line2 As Line, Optional tolerance As Integer? = Nothing) As Boolean
            Return Vector.AreCollinear(line1.Vector, line2.Vector, tolerance)
        End Function

        ''' <summary>
        ''' Check in 3D if two lines = segments are parallel.
        ''' </summary>
        ''' It's not a check of parallelism
        ''' <param name="line1"> First Line </param>
        ''' <param name="line2"> Second Line </param>
        ''' <param name="tolerance">Set the Tolerance (evaluate the vectors proper coordinates), 10^DefaultTolerance is used by default. </param>
        ''' <returns> Result of check as Boolean </returns>
        Public Shared Function AreParallel(line1 As Line, line2 As Line, Optional tolerance As Double? = Nothing) As Boolean
            Return Vector.AreCollinear(line1.Vector, line2.Vector, tolerance)
        End Function

        Public Overloads Function ToString() As String
            Return Vector.ToString()
        End Function

        Public Function Clone() As Line
            Return Me.MemberwiseClone
        End Function
    End Class
