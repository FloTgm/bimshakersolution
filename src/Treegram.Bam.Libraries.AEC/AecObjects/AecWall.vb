Imports M4D.Treegram.Core.Entities
Imports SharpDX
Imports Treegram.GeomLibrary
Imports M4D.Treegram.Core.Extensions.Entities


Public Class AecWall
    Inherits AecObject

    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub


    Private _wallTypo As WallTypo?
    Public Property WallTypo As WallTypo
        Get
            If _wallTypo Is Nothing Then
                _wallTypo = System.[Enum].Parse(GetType(WallTypo), Me.Metaobject.GetAttribute("ResAnaWallTypo", True).Value)
            End If
            Return _wallTypo
        End Get
        Set(value As WallTypo)
            _wallTypo = value
        End Set
    End Property

    Private _axis2D As Curve2D
    Public Property Axis As Curve2D
        Get
            If Me.AxisExtension IsNot Nothing AndAlso _axis2D Is Nothing Then
                Dim axis2d = New Curve2D
                If axis2d.Read(Me.AxisExtension) Then
                    _axis2D = axis2d
                End If
            End If
            Return _axis2D
        End Get
        Set(value As Curve2D)
            _axis2D = value
        End Set
    End Property

    Private _axis As Curve3D
    Public Property Axis(z As Double) As Curve3D
        Get
            If Me.AxisExtension IsNot Nothing AndAlso (_axis Is Nothing OrElse Math.Abs(_axis.Points.First.Z - z) > 0.0001) Then
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

End Class







