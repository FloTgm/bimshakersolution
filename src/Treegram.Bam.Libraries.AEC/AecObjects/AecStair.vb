Imports M4D.Treegram.Core.Entities
Imports M4D.Treegram.Core.Extensions.Entities


Public Class AecStair
    Inherits AecObject

    'THIS CLASS SHOULD BE USED ONLY WITH A STAIRGROUP METAOBJECT !!!!!!!!!!
    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub

    Public ReadOnly Property StairType As StairTypo
        Get
            Dim stairTypeSt = Me.Metaobject.GetAttribute("StairType", True)?.Value
            If stairTypeSt IsNot Nothing AndAlso System.Enum.GetNames(GetType(StairTypo)).Contains(stairTypeSt) Then
                Return System.Enum.Parse(GetType(StairTypo), stairTypeSt)
            Else
                Return StairTypo.Nc
            End If
        End Get
    End Property


    'Private _Stairs As List(Of MetaObject)
    'Public ReadOnly Property StairsComposingStairGroup As List(Of MetaObject)
    '    Get
    '        If _Stairs Is Nothing Then
    '            _Stairs = Me.Metaobject.GetChildren("IfcStair").ToList
    '            If _Stairs.Count = 0 Then
    '                _Stairs = Me.Metaobject.GetChildren("IfcSpace").ToList
    '            End If
    '        End If
    '        Return _Stairs
    '    End Get
    'End Property

    Public ReadOnly Property DistributedSpaces As List(Of AecSpace)
        Get
            Dim myList As New List(Of AecSpace)
            For Each spaceTgm In Me.GetWritableTgmObj.GetParents.ToList
                If spaceTgm.GetTgmType = "SpaceByTgm" Or spaceTgm.GetTgmType = "IfcSpace" Then
                    myList.Add(New AecSpace(spaceTgm))
                End If
            Next
            Return myList
        End Get
    End Property

    Public ReadOnly Property DistributedShells As List(Of AecSpace)
        Get
            Dim myList As New List(Of AecSpace)
            For Each shellTgm In Me.GetWritableTgmObj.GetParents.ToList
                'If shellTgm.GetTgmType = "ShellByTgm" Or shellTgm.GetTgmType = "IfcSpace" Then
                If shellTgm.GetAttribute("IsSpace", True)?.Value = True Then
                    myList.Add(New AecSpace(shellTgm))
                End If
            Next
            Return myList
        End Get
    End Property




End Class

