Imports M4D.Treegram.Core.Entities
Imports Treegram.ConstructionManagement
Imports M4D.Treegram.Core.Extensions.Entities

Public Class RvtFileObject
    Inherits FileObject
    Public Sub New(fileTgm As MetaObject)
        MyBase.New(fileTgm)
    End Sub

    Public ReadOnly Property Storeys As List(Of MetaObject)
        Get
        End Get
    End Property
    Public ReadOnly Property Worksets As List(Of MetaObject)
        Get
        End Get
    End Property


End Class
