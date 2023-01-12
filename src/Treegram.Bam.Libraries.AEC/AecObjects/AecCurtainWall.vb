Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports SharpDX
Imports Treegram.GeomLibrary

Public Class AecCurtainWall
    Inherits AecObject

    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub

    Private _axis2d As Curve2D
    Public Property Axis As Curve2D
        Get
            If Me.AxisExtension IsNot Nothing AndAlso _axis2d Is Nothing Then
                Dim axis2d = New Curve2D
                If axis2d.Read(Me.AxisExtension) Then
                    _axis2d = axis2d
                End If
            End If
            Return _axis2d
        End Get
        Set(value As Curve2D)
            _axis2d = value
        End Set
    End Property

    Private _axis As Curve3D
    Public Property Axis(z As Double) As Curve3D
        Get
            If Me.AxisExtension IsNot Nothing AndAlso (_axis Is Nothing OrElse _axis.Points.First.Z <> z) Then
                Dim axis2d = New Curve2D
                If axis2d.Read(Me.AxisExtension) Then
                    Dim point3dList As New List(Of Vector3)
                    For Each pt In axis2d.Points
                        point3dList.Add(New Vector3(pt.X, pt.Y, z))
                    Next
                    _axis = New Curve3D(point3dList)
                End If
            End If
            Return _axis
        End Get
        Set(value As Curve3D)
            _axis = value
        End Set
    End Property

    ''' <summary>Get the corresponding opening element according to Ifc structure</summary>
    Public ReadOnly Property Reservation As AecObject
        Get
            Dim openingTgm = Me.Metaobject.GetParents("IfcOpeningElement").FirstOrDefault
            If openingTgm Is Nothing Then
                Return Nothing
            Else
                Return New AecObject(openingTgm)
            End If
        End Get
    End Property
End Class






