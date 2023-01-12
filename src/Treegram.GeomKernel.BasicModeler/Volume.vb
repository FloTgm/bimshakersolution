''' <summary>
''' Represents a Volume
''' </summary>
Public Class Volume
    ''' <summary>
    ''' Types of Volume
    ''' </summary>
    Public Enum VolumeType
        Parallelepiped
        Cylinder
    End Enum

    ''' <summary>
    ''' List of Points of the Volume
    ''' </summary>
    ''' <returns></returns>
    Public Property Points As List(Of Point)

    ''' <summary>
    ''' Type of the Volume
    ''' </summary>
    ''' <returns></returns>
    Public Property Type As VolumeType

    ''' <summary>
    ''' Creates a new Volume by a list of Points
    ''' </summary>
    ''' <param name="points"></param>
    ''' <param name="type"></param>
    Public Sub New(points As List(Of Point), type As VolumeType)
        Me.Points = points
        Me.Type = type
    End Sub

    ''' <summary>
    ''' Volume measured
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Volume As Double
        Get
            Select Case Type
                Case VolumeType.Parallelepiped
                    Dim longueur As Double = Distance(Points(0), Points(1))
                    Dim largeur As Double = Distance(Points(0), Points(2))
                    Dim hauteur As Double = Distance(Points(0), Points(3))
                    Return Math.Round(RoundToSignificantDigits(longueur * largeur * hauteur, RealPrecision), -DefaultTolerance - 6)
                Case VolumeType.Cylinder
                    Dim rayon As Double = Distance(Points(0), Points(1))
                    Dim hauteur As Double = Distance(Points(0), Points(2))
                    Return Math.Round(RoundToSignificantDigits(Math.PI * rayon * rayon * hauteur, RealPrecision), -DefaultTolerance - 6)
                Case Else
                    Throw New Exception("define volume type before checking the volume")
            End Select
        End Get
    End Property

    Friend ReadOnly Property InternVolume As Double
        Get
            Select Case Type
                Case VolumeType.Parallelepiped
                    Dim longueur As Double = Distance(Points(0), Points(1))
                    Dim largeur As Double = Distance(Points(0), Points(2))
                    Dim hauteur As Double = Distance(Points(0), Points(3))
                    Return RoundToSignificantDigits(longueur * largeur * hauteur, RealPrecision)
                Case VolumeType.Cylinder
                    Dim rayon As Double = Distance(Points(0), Points(1))
                    Dim hauteur As Double = Distance(Points(0), Points(2))
                    Return RoundToSignificantDigits(Math.PI * rayon * rayon * hauteur, RealPrecision)
                Case Else
                    Throw New Exception("define volume type before checking the volume")
            End Select
        End Get
    End Property

End Class
