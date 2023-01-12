Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities
Imports Treegram.ConstructionManagement
Imports Treegram.GeomLibrary

Public Class AecOpening
    Inherits AecObject


    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
    End Sub

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
    Public ReadOnly Property HasReservation As Boolean
        Get
            Return Me.Metaobject.GetAttribute("HasReservation", True)?.Value
        End Get
    End Property
    Public ReadOnly Property ReservationProfileIsCorrect As Boolean
        Get
            Return Me.Metaobject.GetAttribute("ReservationProfileIsCorrect", True)?.Value
        End Get
    End Property


    Public ReadOnly Property GlazedAreaByTgm As Double
        Get
            Return Me.Metaobject.GetAttribute("GlazedAreaByTgm", True).Value
        End Get
    End Property
    Public ReadOnly Property Orientation As GeographicOrientation
        Get
            Dim orient = Me.Metaobject.GetAttribute("OrientationByTgm", True)?.Value
            If orient IsNot Nothing AndAlso System.Enum.GetNames(GetType(GeographicOrientation)).Contains(orient.ToString) Then
                Return System.[Enum].Parse(GetType(GeographicOrientation), orient.ToString)
            Else
                Return Nothing
            End If
        End Get
    End Property

    Dim _overallProfileExtTgm As MetaObject
    Public ReadOnly Property OverallProfileExtension As MetaObject
        Get
            If _overallProfileExtTgm Is Nothing Then
                _overallProfileExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 17)) = " - OverallProfile")
            End If
            Return _overallProfileExtTgm
        End Get
    End Property
    Private _overallProfile As Profile3D
    Public ReadOnly Property OverallProfile As Profile3D
        Get
            If _overallProfile Is Nothing AndAlso Me.OverallProfileExtension IsNot Nothing Then
                Dim profileAtt = Me.OverallProfileExtension.GetAttribute("OverallProfile", True)
                If profileAtt IsNot Nothing Then
                    Dim myProfile As New Profile3D
                    If myProfile.Read(profileAtt) Then
                        _overallProfile = myProfile
                    End If
                End If
            End If
            Return _overallProfile
        End Get
    End Property

    Dim _openingProfileExtTgm As MetaObject
    Public ReadOnly Property OpeningProfileExtension As MetaObject
        Get
            If _openingProfileExtTgm Is Nothing Then
                _openingProfileExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 17)) = " - OpeningProfile")
            End If
            Return _openingProfileExtTgm
        End Get
    End Property
    Private _openingProfile As Profile3D
    Public ReadOnly Property OpeningProfile As Profile3D
        Get
            If _openingProfile Is Nothing AndAlso Me.OpeningProfileExtension IsNot Nothing Then
                Dim profileAtt = Me.OpeningProfileExtension.GetAttribute("OpeningProfile", True)
                If profileAtt IsNot Nothing Then
                    Dim myProfile As New Profile3D
                    If myProfile.Read(profileAtt) Then
                        _openingProfile = myProfile
                    End If
                End If
            End If
            Return _openingProfile
        End Get
    End Property

    Dim _glazingProfileExtTgm As MetaObject
    Public ReadOnly Property GlazingProfileExtension As MetaObject
        Get
            If _glazingProfileExtTgm Is Nothing Then
                _glazingProfileExtTgm = Me.Metaobject.Extensions.FirstOrDefault(Function(o) o.Name.Substring(Math.Max(0, o.Name.Length - 10)) = " - Glazing")
            End If
            Return _glazingProfileExtTgm
        End Get
    End Property

    Private _glazingProfile As List(Of Profile3D)
    Public ReadOnly Property GlazingProfile As List(Of Profile3D)
        Get
            If _glazingProfile Is Nothing AndAlso Me.GlazingProfileExtension IsNot Nothing Then
                Dim glazProfileAttributes = Me.GlazingProfileExtension.Attributes.Where(Function(o) o.Name.Split(".").First = "GlazingProfile").ToList
                For Each profileAtt In glazProfileAttributes
                    Dim myProfile As New Profile3D
                    If myProfile.Read(profileAtt) Then
                        If _glazingProfile Is Nothing Then
                            _glazingProfile = New List(Of Profile3D) From {myProfile}
                        Else
                            _glazingProfile.Add(myProfile)
                        End If
                    End If
                Next
            End If
            Return _glazingProfile
        End Get
    End Property

    ''' <summary>Get the host wall according to Tgm analysis</summary>
    Public ReadOnly Property WallOwner As AecWall
        Get
            Dim wallTgm As MetaObject = Nothing
            For Each parent In Me.Metaobject.GetParents
                If parent.GetTgmType = ProjectReference.Constants.overlappingObjType Then
                    wallTgm = parent
                ElseIf parent.Extend IsNot Nothing AndAlso parent.Extend.GetAttribute("IsVerticalSeparator", True)?.Value = True Then
                    wallTgm = parent.Extend
                End If
            Next
            If wallTgm Is Nothing Then
                Return Nothing
            Else
                Return New AecWall(wallTgm)
            End If
        End Get
    End Property

    Public ReadOnly Property DistributedSpaces As List(Of AecSpace)
        Get
            Dim myList As New List(Of AecSpace)
            For Each spaceTgm In Me.Metaobject.GetParents.ToList
                If spaceTgm.GetTgmType = "SpaceByTgm" Then
                    myList.Add(New AecSpace(spaceTgm))
                End If
            Next
            Return myList
        End Get
    End Property
    Public ReadOnly Property DistributedShells As List(Of AecSpace)
        Get
            Dim myList As New List(Of AecSpace)
            For Each spaceTgm In Me.Metaobject.GetParents.ToList
                'If spaceTgm.GetTgmType = "ShellByTgm" Or spaceTgm.GetTgmType = "IfcSpace" Then
                If spaceTgm.GetAttribute("IsSpace", True)?.Value = True Then
                    myList.Add(New AecSpace(spaceTgm))
                End If
            Next
            Return myList
        End Get
    End Property

    Public Property FromSpace As AecSpace
    Public Property ToSpace As AecSpace


    Private _openEltInsertPt As Point3D
    Public ReadOnly Property IfcOpeningEltInsertPoint As Point3D
        Get
            If _openEltInsertPt Is Nothing Then

                '---BY IFC INFOS
                'insertpoint
                Dim localPlacement = Me.Metaobject.GetAttribute("LocalPlacement")
                If localPlacement Is Nothing Then Return Nothing
                Dim location = localPlacement.GetAttribute("Location")
                If location Is Nothing Then Return Nothing
                Dim xpoint = CDbl(location.GetAttribute("X").Value)
                Dim ypoint = CDbl(location.GetAttribute("Y").Value)
                Dim zpoint = CDbl(location.GetAttribute("Z").Value)
                Dim insertPoint As New Point3D(xpoint, ypoint, zpoint)

                ''translation en X
                'Dim AxisX = LocalPlacement.GetAttribute("Axis X")
                'Dim XaxisX = CDbl(AxisX.GetAttribute("X").Value)
                'Dim YaxisX = CDbl(AxisX.GetAttribute("Y").Value)
                'Dim VectXAxis As new Treegram.GeomKernel.BasicModeler.Vector(XaxisX, YaxisX, 0.0)

                ''translation en Y
                'Dim AxisY = LocalPlacement.GetAttribute("Axis Y")
                'Dim XaxisY = CDbl(AxisY.GetAttribute("X").Value)
                'Dim YaxisY = CDbl(AxisY.GetAttribute("Y").Value)
                'Dim VectYAxis As new Treegram.GeomKernel.BasicModeler.Vector(XaxisY, YaxisY, 0.0)

                ''translation en X
                'Dim translatX = VectXAxis.GetPointAtDistance(0.2, doorpoint)

                ''translation en Y
                '_OpenEltCenterPt = VectYAxis.GetPointAtDistance(0.02, translatX)
                _openEltInsertPt = insertPoint
            End If
            Return _openEltInsertPt
        End Get
    End Property



End Class

