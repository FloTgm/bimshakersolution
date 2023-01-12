''' <summary>
''' Use for "treegram 3D triangle to face" algorithm
''' </summary>
Public Class Triangle

    Public Property P1 As Point

    Public Property P2 As Point

    Public Property P3 As Point

    Public ReadOnly Property Normal As Vector
        Get
            Return Vector.CrossProduct(New Vector(P1, P2), New Vector(P1, P3))
        End Get
    End Property

    Public ReadOnly Property Points As List(Of Point)
        Get
            Return New List(Of Point)() From {P1, P2, P3}
        End Get
    End Property

    Public Sub New(p1 As Point, p2 As Point, p3 As Point)
        Me.P1 = p1
        Me.P2 = p2
        Me.P3 = p3
    End Sub

End Class
