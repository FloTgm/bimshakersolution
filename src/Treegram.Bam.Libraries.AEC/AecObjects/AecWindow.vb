Imports M4D.Treegram.Core.Entities

Public Class AecWindow
    Inherits AecOpening

    Public Sub New(objTgm As MetaObject)
        MyBase.New(objTgm)
        If objTgm Is Nothing Then
            Throw New Exception("This kind of analysis object must be defined by a metaobject")
        End If
    End Sub

    Dim _windowTypo As WindowTypo?
    Public Property WindowTypo As WindowTypo?
        Get
            If _windowTypo Is Nothing Then
                Dim typoAttr = Me.Metaobject.GetAttribute("WindowTypeByTgm", True)?.Value
                If typoAttr IsNot Nothing Then
                    _windowTypo = [Enum].Parse(GetType(WindowTypo), typoAttr)
                End If

            End If
            Return _windowTypo
        End Get
        Set(value As WindowTypo?)
            _windowTypo = value
        End Set
    End Property
End Class



