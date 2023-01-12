Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities

Public Class AecSpace
    Inherits AecObject

    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub

    Private _spaceTypo As SpaceTypo?
    Public Property SpaceTypo As SpaceTypo
        Get
            If _spaceTypo Is Nothing Then
                Dim typeAtt = Me.Metaobject.GetAttribute("Type", True)
                If typeAtt IsNot Nothing AndAlso System.Enum.GetNames(GetType(SpaceTypo)).Contains(typeAtt.Value.ToString) Then
                    _spaceTypo = System.[Enum].Parse(GetType(SpaceTypo), typeAtt.Value.ToString)
                Else
                    _spaceTypo = SpaceTypo.Inconnu
                End If
            End If
            Return _spaceTypo
        End Get
        Set(value As SpaceTypo)
            _spaceTypo = value
        End Set
    End Property


    Public ReadOnly Property IsCorridor As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsCorridorSpace", True)?.Value
        End Get
    End Property
    Public ReadOnly Property IsStairCase As Boolean
        Get
            If Me.Metaobject.GetAttribute("StairType", True)?.Value = StairTypo.Commun.ToString Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property IsDuplex As Boolean
        Get
            If Me.Metaobject.GetAttribute("StairType", True)?.Value = StairTypo.Privé.ToString Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


    Public ReadOnly Property Source As String
        Get
            If Me.Metaobject.GetAttribute("Source") IsNot Nothing Then
                Return Me.Metaobject.GetAttribute("Source").Value
            Else
                Return Nothing
            End If
        End Get
    End Property
    Public ReadOnly Property Support As String
        Get
            If Me.Metaobject.GetAttribute("Support") IsNot Nothing Then
                Return Me.Metaobject.GetAttribute("Support").Value
            Else
                Return Nothing
            End If
        End Get
    End Property



    ''' <summary>Get the doors connected to the space </summary>
    Private _doors As List(Of AecDoor) = Nothing
    Public Property Doors As List(Of AecDoor)
        Get
            'If _Doors Is Nothing Then
            _doors = New List(Of AecDoor)
            For Each childTgm In Me.GetWritableTgmObj.GetChildren
                Dim potentialDoor As New AecDoor(childTgm)
                If potentialDoor.Category = Category.Door Then
                    _doors.Add(potentialDoor)
                ElseIf childTgm.GetTgmType = "IfcDoor" Or childTgm.GetAttribute("RevitCategory")?.Value = "Portes" Then 'Old project case...
                    _doors.Add(potentialDoor)
                End If
            Next
            'End If
            Return _doors
        End Get
        Set(value As List(Of AecDoor))
            _doors = value
        End Set
    End Property

    ''' <summary>Get the stairs connected to the space </summary>
    Private _cStairs As List(Of AecStair) = Nothing
    Public Property CStairs As List(Of AecStair)
        Get
            If _cStairs Is Nothing Then
                Dim childStairsList = Me.GetWritableTgmObj.GetChildren("StairGroup").ToList
                _cStairs = New List(Of AecStair)
                For Each stairTgm In childStairsList
                    _cStairs.Add(New AecStair(stairTgm))
                Next
            End If
            Return _cStairs
        End Get
        Set(value As List(Of AecStair))
            _cStairs = value
        End Set
    End Property
    Private _windows As List(Of AecWindow) = Nothing
    Public Property Windows As List(Of AecWindow)
        Get
            _windows = New List(Of AecWindow)
            For Each childTgm In Me.GetWritableTgmObj.GetChildren
                Dim potentialWindow As New AecWindow(childTgm)
                If potentialWindow.Category = Category.Window Then
                    _windows.Add(potentialWindow)
                ElseIf childTgm.GetTgmType = "IfcWindow" Or childTgm.GetTgmType = "IfcCurtainWall" Then 'Old project case...
                    _windows.Add(potentialWindow)
                End If
            Next
            Return _windows
        End Get
        Set(value As List(Of AecWindow))
            _windows = value
        End Set
    End Property

    Private _verticalReservations As List(Of AecOpening) = Nothing
    Public Property VerticalReservations As List(Of AecOpening)
        Get
            Dim childList As New List(Of MetaObject)
            For Each childTgm In Me.GetWritableTgmObj.GetChildren
                If childTgm.GetTgmType = "IfcOpeningElement" Then
                    childList.Add(childTgm)
                End If
            Next
            _verticalReservations = New List(Of AecOpening)
            For Each objTgm In childList
                _verticalReservations.Add(New AecOpening(objTgm))
            Next
            Return _verticalReservations
        End Get
        Set(value As List(Of AecOpening))
            _verticalReservations = value
        End Set
    End Property

    Public ReadOnly Property AllOpenings As List(Of AecOpening)
        Get
            Dim childOpeningsList As New List(Of AecOpening)
            For Each childTgm In Me.GetWritableTgmObj.GetChildren
                Dim potentialOpening As New AecOpening(childTgm)
                If potentialOpening.IsOpenable Then
                    childOpeningsList.Add(potentialOpening)
                End If
            Next
            Return childOpeningsList
        End Get
    End Property

    Public Property AlreadyDone As Boolean = False

    Public Property ConnectedSpaces As List(Of AecSpace)


End Class
