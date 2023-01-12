Imports M4D.Treegram.Core.Entities

Public Class AecDoor
    Inherits AecOpening


    Dim _doorTypo As DoorTypo?
    Dim _doorTypoByWall, _doorTypoByShell, _doorTypoBySpace As DoorTypo?


    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of object must be defined by a metaobject")
        End If
    End Sub

    Public ReadOnly Property DoorType As DoorTypo?
        Get
            If _doorTypo Is Nothing Then
                Dim typoAttr = Me.Metaobject.GetAttribute("TypeByTgm", True)?.Value
                If typoAttr Is Nothing Then
                    typoAttr = Me.Metaobject.GetAttribute("DoorTypeByTgm", True)?.Value 'OLD ONE !
                End If
                If typoAttr IsNot Nothing Then
                    _doorTypo = [Enum].Parse(GetType(DoorTypo), typoAttr)
                End If
            End If
            Return _doorTypo
        End Get
    End Property

    Public Property DoorTypoByWall As DoorTypo?
        Get
            If _doorTypoByWall Is Nothing Then
                Dim typoAttr = Me.Metaobject.GetAttribute("AnaDoorTypoByWall", True)?.Value
                If typoAttr IsNot Nothing Then
                    _doorTypoByWall = [Enum].Parse(GetType(DoorTypo), typoAttr)
                End If

            End If
            Return _doorTypoByWall
        End Get
        Set(value As DoorTypo?)
            _doorTypoByWall = value
        End Set
    End Property
    Public Property DoorTypoBySpace As DoorTypo?
        Get
            If _doorTypoBySpace Is Nothing Then
                _doorTypoBySpace = [Enum].Parse(GetType(DoorTypo), Me.Metaobject.GetAttribute("AnaDoorTypoBySpace", True)?.Value)
            End If
            Return _doorTypoBySpace
        End Get
        Set(value As DoorTypo?)
            _doorTypoBySpace = value
        End Set
    End Property

    Public Property DoorTypoByShell As DoorTypo?
        Get
            If _doorTypoByShell Is Nothing Then
                _doorTypoByShell = [Enum].Parse(GetType(DoorTypo), Me.Metaobject.GetAttribute("AnaDoorTypoByShell", True)?.Value)
            End If
            Return _doorTypoByShell
        End Get
        Set(value As DoorTypo?)
            _doorTypoByShell = value
        End Set
    End Property
End Class




