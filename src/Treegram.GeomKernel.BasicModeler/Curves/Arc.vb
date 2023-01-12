''' <summary>
''' Represents a 2D-Circle Arc, angles are trigonometric angles in [0;2*Pi[
''' </summary>
Public Class Arc : Inherits Curve

    ''' <summary>
    ''' CenterPoint of the Arc
    ''' </summary>
    Private _mCenterPoint As Point

    ''' <summary>
    ''' Get or Set centerPoint of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CenterPoint As Point
        Get
            Return Me._mCenterPoint
        End Get
    End Property

    ''' <summary>
    ''' Radius of the Arc
    ''' </summary>
    Private _mRadius As Double

    ''' <summary>
    ''' Get or Set the radius of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Radius As Double
        Get
            Return Me._mRadius
        End Get
    End Property

    ''' <summary>
    ''' StartAngle of the Arc
    ''' </summary>
    Private _mStartAngle As Double

    ''' <summary>
    ''' Get or Set the startAngle of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property StartAngle As Double
        Get
            Return Me._mStartAngle
        End Get
    End Property

    ''' <summary>
    ''' EndAngle of the Arc
    ''' </summary>
    Private _mEndAngle As Double

    ''' <summary>
    ''' Get or Set the endAngle of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property EndAngle As Double
        Get
            Return Me._mEndAngle
        End Get
    End Property

    ''' <summary>
    ''' CounterClockWise of the Arc
    ''' </summary>
    Private _mIsCounterClockWise As Boolean

    ''' <summary>
    ''' Get or Set the counterClockWise of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public Property IsCounterClockWise As Boolean
        Get
            Return Me._mIsCounterClockWise
        End Get
        Set(isCounterClockWise As Boolean)
            Me._mIsCounterClockWise = isCounterClockWise
        End Set
    End Property

    ''' <summary>
    ''' Get the length of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public Overrides ReadOnly Property Length As Double
        Get
            Dim angle As New Double
            If IsCounterClockWise Then
                angle = (EndAngle - StartAngle) Mod (2 * Math.PI)
            Else
                angle = (StartAngle - EndAngle) Mod (2 * Math.PI)
            End If
            If angle < 0 Then
                angle += 2 * Math.PI
            End If
            Return Math.Round(RoundToSignificantDigits(Math.Abs(Radius * angle), RealPrecision), -DefaultTolerance)
        End Get
    End Property

    Friend ReadOnly Property InternLength As Double
        Get
            Dim angle As New Double
            If IsCounterClockWise Then
                angle = (EndAngle - StartAngle) Mod (2 * Math.PI)
            Else
                angle = (StartAngle - EndAngle) Mod (2 * Math.PI)
            End If
            If angle < 0 Then
                angle += 2 * Math.PI
            End If
            Return RoundToSignificantDigits(Math.Abs(Radius * angle), RealPrecision)
        End Get
    End Property

    ''' <summary>
    ''' Get a Point on Arc by angle
    ''' </summary>
    ''' <param name="angle"> Angle </param>
    ''' <returns> Point on Arc </returns>
    Public Function Point(angle As Double) As Point
        Return CenterPoint + New Point(Math.Cos(angle) * Me.Radius, Math.Sin(angle) * Me.Radius)
    End Function

    ''' <summary>
    ''' Get the MidPoint of the Arc
    ''' </summary>
    ''' <returns> MidPoint </returns>
    Public Overrides ReadOnly Property MidPoint As Point
        Get
            Dim angle As New Double
            If IsCounterClockWise Then
                angle = (EndAngle - StartAngle) Mod (2 * Math.PI)
            Else
                angle = (StartAngle - EndAngle) Mod (2 * Math.PI)
            End If
            If angle < 0 Then
                angle += (2 * Math.PI)
            End If
            angle = angle / 2
            If IsCounterClockWise Then
                angle = (StartAngle + angle) Mod (2 * Math.PI)
            Else
                angle = (StartAngle - angle) Mod (2 * Math.PI)
            End If
            Return Point(angle)
        End Get
    End Property

    Private _mStartPoint As Point = Nothing

    ''' <summary>
    ''' Get StartPoint of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Property StartPoint As Point
        Get
            If IsNothing(_mStartPoint) Then
                Return Point(StartAngle)
            Else
                Return _mStartPoint
            End If
        End Get
        Set(value As Point)
            Throw New NotImplementedException
        End Set
    End Property

    Private _mEndPoint As Point = Nothing

    ''' <summary>
    ''' Get EndPoint of the Arc
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Property EndPoint As Point
        Get
            If IsNothing(_mEndPoint) Then
                Return Point(EndAngle)
            Else
                Return _mEndPoint
            End If
        End Get
        Set(value As Point)
            Throw New NotImplementedException()
        End Set
    End Property

    ''' <summary>
    ''' Create new Arc by centerPoint, radius and angles
    ''' </summary>
    ''' <param name="centerPoint"> CenterPoint </param>
    ''' <param name="radius"> Radius </param>
    ''' <param name="startAngle"> StartAngle in radians </param>
    ''' <param name="endAngle"> EndAngle in radians </param>
    ''' <param name="isCounterClockWise"> Trigonomic(true) or hourly(false) </param>
    Public Sub New(centerPoint As Point, radius As Double, startAngle As Double, endAngle As Double, Optional isCounterClockWise As Boolean = True)
        _mCenterPoint = centerPoint
        _mRadius = Math.Round(RoundToSignificantDigits(radius, RealPrecision), -DefaultTolerance)
        Dim stAngle = startAngle Mod (2 * Math.PI)
        If stAngle < 0 Then
            _mStartAngle = RoundToSignificantDigits(stAngle + 2 * Math.PI, RealPrecision)
        Else
            _mStartAngle = RoundToSignificantDigits(stAngle, RealPrecision)
        End If
        Dim edAngle = endAngle Mod (2 * Math.PI)
        If edAngle < 0 Then
            _mEndAngle = RoundToSignificantDigits(edAngle + 2 * Math.PI, RealPrecision)
        Else
            _mEndAngle = RoundToSignificantDigits(edAngle, RealPrecision)
        End If
        Me.IsCounterClockWise = isCounterClockWise
    End Sub

    ''' <summary>
    ''' Create nesw Arc by centerPoint, startPoint and endPoint
    ''' </summary>
    ''' <param name="centerPoint"> CenterPoint </param>
    ''' <param name="startPoint"> StartPoint </param>
    ''' <param name="endPoint"> EndPoint </param>
    ''' <param name="isCounterClockWise"> Trigonomic(true) or hourly(false) </param>
    Public Sub New(centerPoint As Point, startPoint As Point, endPoint As Point, Optional isCounterClockWise As Boolean = True)
        _mCenterPoint = centerPoint
        Me._mStartPoint = startPoint
        Me._mEndPoint = endPoint
        _mRadius = Distance(centerPoint, startPoint)
        Me.IsCounterClockWise = isCounterClockWise
        Dim stAngle = New Vector(centerPoint, startPoint).Angle(CType(New Point(1, 0), Vector))
        If stAngle < 0 Then
            Me._mStartAngle = RoundToSignificantDigits(stAngle + 2 * Math.PI, RealPrecision)
        Else
            Me._mStartAngle = RoundToSignificantDigits(stAngle, RealPrecision)
        End If
        Dim edAngle = New Vector(centerPoint, endPoint).Angle(CType(New Point(1, 0), Vector))
        If edAngle < 0 Then
            Me._mEndAngle = RoundToSignificantDigits(edAngle + 2 * Math.PI, RealPrecision)
        Else
            Me._mEndAngle = RoundToSignificantDigits(edAngle, RealPrecision)
        End If
    End Sub

    ''' <summary>
    ''' 'Create new Arc by startPoint, endPoint and any point on Arc betrween the two
    ''' </summary>
    ''' <param name="startPoint"> StartPoint </param>
    ''' <param name="otherPoint"> Any Point on Arc between Start and End </param>
    ''' <param name="endPoint"> EndPoint </param>
    Public Sub New(startPoint As Point, otherPoint As Point, endPoint As Point)
        Dim p1, p2, p3, centerPoint As Point
        Dim xC, yC As Double
        Dim moyenne = False
        If startPoint.Y = otherPoint.Y Then 'startPoint/otherPoint horizontal
            If startPoint.X = endPoint.X OrElse otherPoint.X = endPoint.X Then 'start/end vertical ou other/end vertical
                p1 = startPoint
                p2 = otherPoint
                p3 = endPoint
                moyenne = True
            Else
                p1 = otherPoint
                p2 = endPoint
                p3 = startPoint
            End If
        Else
            If otherPoint.Y = endPoint.Y Then 'other/end horizontal
                If startPoint.X = otherPoint.X OrElse startPoint.X = endPoint.X Then 'start/other ou start/end vertical
                    p1 = startPoint
                    p2 = otherPoint
                    p3 = endPoint
                    moyenne = True
                Else
                    p1 = endPoint
                    p2 = startPoint
                    p3 = otherPoint
                End If
            Else
                p1 = startPoint
                p2 = otherPoint
                p3 = endPoint
            End If
        End If

        If moyenne Then
            If p1.X = p2.X Then
                xC = (p1.X + p3.X) / 2
            Else
                xC = (p1.X + p2.X) / 2
            End If
            If p1.Y = p2.Y Then
                yC = (p1.Y + p3.Y) / 2
            Else
                yC = (p1.Y + p2.Y) / 2
            End If
        Else
            'xC = ((P3.X * P3.X - P2.X * P2.X + P3.Y * P3.Y - P2.Y * P2.Y) / (2 * P3.Y - 2 * P2.Y) - (P2.X * P2.X - P1.X * P1.X + P2.Y * P2.Y - P1.Y * P1.Y) / (2 * P2.Y - 2 * P1.Y)) /
            '    ((P3.X - P2.X) / (P3.Y - P2.Y) - (P2.X - P1.X) / (P2.Y - P1.Y))
            'yC = -(P2.X - P1.X) / (P2.Y - P1.Y) * xC + (P2.X * P2.X - P1.X * P1.X + P2.Y * P2.Y - P1.Y * P1.Y) / (2 * P2.Y - 2 * P1.Y)
            xC = ((p3.X * p3.X - p2.X * p2.X) / (p3.Y - p2.Y) - (p1.X * p1.X - p2.X * p2.X) / (p1.Y - p2.Y) + p3.Y - p1.Y) /
                     (2 * ((p3.X - p2.X) / (p3.Y - p2.Y)) - 2 * ((p1.X - p2.X) / (p1.Y - p2.Y)))
            yC = -(p2.X - p1.X) / (p2.Y - p1.Y) * xC + (p2.X * p2.X - p1.X * p1.X) / (2 * p2.Y - 2 * p1.Y) + p1.Y / 2 + p2.Y / 2
        End If
        If Double.IsInfinity(xC) OrElse Double.IsInfinity(yC) Then
            Throw New Exception("Calculated center is at infinity. Points should not be aligned.")
        End If
        centerPoint = New Point(xC, yC)

        Me._mCenterPoint = centerPoint
        Me._mStartPoint = startPoint
        Me._mEndPoint = endPoint
        'Application.PRECISION = 10
        Me._mRadius = Distance(centerPoint, startPoint)
        Dim stAngle = New Vector(centerPoint, startPoint).Angle(CType(New Point(1, 0), Vector))
        If stAngle < 0 Then
            Me._mStartAngle = RoundToSignificantDigits(stAngle + 2 * Math.PI, RealPrecision)
        Else
            Me._mStartAngle = RoundToSignificantDigits(stAngle, RealPrecision)
        End If
        Dim edAngle = New Vector(centerPoint, endPoint).Angle(CType(New Point(1, 0), Vector))
        If edAngle < 0 Then
            Me._mEndAngle = RoundToSignificantDigits(edAngle + 2 * Math.PI, RealPrecision)
        Else
            Me._mEndAngle = RoundToSignificantDigits(edAngle, RealPrecision)
        End If

            If Vector.CrossProduct(New Vector(startPoint, otherPoint), New Vector(otherPoint, endPoint)).Z > 0 Then
                Me.IsCounterClockWise = True
            Else
                Me.IsCounterClockWise = False
            End If
        End Sub

    End Class
