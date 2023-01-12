Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities

Public Class AecApartment
    Inherits AecObject

    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub

    Private _spaces As List(Of AecSpace)
    Public Property ComposingSpaces As List(Of AecSpace)
        Get
            If _spaces Is Nothing Then
                _spaces = Me.Metaobject.GetChildren("SpaceByTgm").ToList.Select(Function(o) New AecSpace(o)).ToList
            End If
            Return _spaces
        End Get
        Set(ByVal value As List(Of AecSpace))
            _spaces = value
        End Set
    End Property

    Private _shells As List(Of AecSpace)
    Public Property ComposingShells As List(Of AecSpace)
        Get
            If _shells Is Nothing Then
                'ATTENTION AVEC LES SpaceByTgm !!!!!!!!!!!!!!!!!!!!!
                _shells = Me.Metaobject.GetChildren(,, RelationType.Decomposition).Select(Function(o) New AecSpace(o)).Where(Function(l) l.IsSpace).ToList
            End If
            Return _shells
        End Get
        Set(ByVal value As List(Of AecSpace))
            _shells = value
        End Set
    End Property
    Public ReadOnly Property IsDuplex As Boolean
        Get
            Return Me.Metaobject.GetAttribute("IsDuplex", True).Value
        End Get
    End Property
End Class